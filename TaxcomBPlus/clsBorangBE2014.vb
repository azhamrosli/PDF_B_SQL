Imports iTextSharp.text.pdf
Imports System.Data.OleDb
Imports System.IO

Public Class clsBorangBE2014
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
        Dim dr As OleDbDataReader = Nothing
        Dim dr2 As OleDbDataReader = Nothing
        Dim prmOledb(0) As OleDbParameter
        Dim strHWIC As String = ""
        Dim strArray(2) As String


        Try
            prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page1[0]."


            ' ==== Master Data ==== "
            ds = datHandler.GetData("select tp_name , (tp_ref_no_prefix + tp_ref_no1 + tp_ref_no2 + tp_ref_no3)," _
                        & " (tp_ic_new_1 + tp_ic_new_2 + tp_ic_new_3)," _
                        & " tp_passport_no," _
                        & " tp_country, tp_gender, tp_status, tp_date_marriage," _
                        & " tp_date_divorce, tp_type_assessment," _
                        & " tp_hw_name," _
                        & " (tp_hw_ic_new1 + tp_hw_ic_new2 + tp_hw_ic_new3)," _
                        & " tp_hw_passport_no, tp_assessmenton," _
                        & " (tp_curr_add_line1 + ', ' + tp_curr_add_line2 + ', ' + tp_curr_add_line3 + ', '+ tp_curr_postcode + ' ' + tp_curr_city + ', ' + tp_curr_state)" _
                        & " from taxp_profile where (tp_ref_no1 + tp_ref_no2 + tp_ref_no3)=?", prmOledb)

            dr = datHandler.GetDataReader("select tp_last_passport_no, TP_WORKER_APPROVEDATE, TP_COM_ADD_STATUS from taxp_profile2 where" _
                        & " tp_ref_no= '" & pdfForm.GetRefNo & "'")

            If ds.Tables(0).Rows.Count > 0 Then
                '========= HEADER =========
                '--- Name ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(0).ToString()) Then
                        pdfFormFields.SetField(pdfFieldPath & "H1[0]", ds.Tables(0).Rows(0).Item(0).ToString.ToUpper)
                    End If
                End If

                '--- Address ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(14)) Then
                    strArray = SplitText(FormatAddress(ds.Tables(0).Rows(0).Item(14).ToString()), 38)
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(14).ToString()) Then
                        pdfFormFields.SetField(pdfFieldPath & "H2[0]", strArray(0).ToUpper)
                        pdfFormFields.SetField(pdfFieldPath & "H3[0]", strArray(1).ToUpper)
                        pdfFormFields.SetField(pdfFieldPath & "H4[0]", strArray(2).ToUpper)
                    End If
                End If

                '--- Date ---
                If Not String.IsNullOrEmpty(pdfForm.GetDeclarationDate) Then
                    pdfFormFields.SetField(pdfFieldPath & "H5[0]", FormatDeclarationDate(pdfForm.GetDeclarationDate))
                End If



                '========= BASIC INFORMATION =========

                '--- 1 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(0).ToString()) Then
                        pdfFormFields.SetField(pdfFieldPath & "MA1[0]", ds.Tables(0).Rows(0).Item(0).ToString.ToUpper)
                    End If
                End If

                '--- 2 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(1)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(1).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "MA2[0]", ds.Tables(0).Rows(0).Item(1).ToString)
                    End If
                End If

                '--- 3 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(2)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(2).ToString) Then
                        'dannylee 2012su2.1
                        If Not Trim(ds.Tables(0).Rows(0).Item(2).ToString) = "" Then
                            pdfFormFields.SetField(pdfFieldPath & "MA3[0]", FormatICNumber(ds.Tables(0).Rows(0).Item(2).ToString))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "MA3[0]", "---")
                        End If
                        'end
                    End If
                End If

                '--- 4 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(3)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(3).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "MA4[0]", ds.Tables(0).Rows(0).Item(3).ToString)
                    End If
                End If

                '--- 5 ---
                If dr.Read() Then
                    If Not IsDBNull(dr("tp_last_passport_no")) Then
                        If Not String.IsNullOrEmpty(dr("tp_last_passport_no").ToString) Then
                            pdfFormFields.SetField(pdfFieldPath & "MA5[0]", dr("tp_last_passport_no").ToString)
                        End If
                    End If
                End If
                dr.Close()


                '========= SECTION A =========

                '--- A1 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(4)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(4).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A1[0]", Mid(CStr((ds.Tables(0).Rows(0).Item(4).ToString)), 1, 1) & Space(3) & Mid(CStr((ds.Tables(0).Rows(0).Item(4).ToString)), 2, 1))
                    End If
                End If

                '--- A2 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(5)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(5).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A2[0]", ds.Tables(0).Rows(0).Item(5).ToString)
                    End If
                End If

                '--- A3 & A4 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(6)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(6).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A3[0]", ds.Tables(0).Rows(0).Item(6).ToString)
                        If ds.Tables(0).Rows(0).Item(6).ToString = "2" Then
                            If Not IsDBNull(ds.Tables(0).Rows(0).Item(7)) Then
                                pdfFormFields.SetField(pdfFieldPath & "A4[0]", FormatDate(ds.Tables(0).Rows(0).Item(7)))
                            End If
                        ElseIf ds.Tables(0).Rows(0).Item(6).ToString = "3" Or ds.Tables(0).Rows(0).Item(6).ToString = "4" Then
                            If Not IsDBNull(ds.Tables(0).Rows(0).Item(8)) Then
                                pdfFormFields.SetField(pdfFieldPath & "A4[0]", FormatDate(ds.Tables(0).Rows(0).Item(8)))
                            End If
                        End If
                    End If
                End If

                '--- A5 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(9)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(9).ToString) Then
                        If ds.Tables(0).Rows(0).Item(9).ToString = "1" Then
                            If ds.Tables(0).Rows(0).Item(13).ToString = "1" Then
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "1")
                            ElseIf ds.Tables(0).Rows(0).Item(13).ToString = "2" Then
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "2")
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "")
                            End If
                        ElseIf ds.Tables(0).Rows(0).Item(9).ToString = "2" Then
                            pdfFormFields.SetField(pdfFieldPath & "A5[0]", "3")
                        ElseIf ds.Tables(0).Rows(0).Item(9).ToString = "3" Then
                            'weihong
                            If ds.Tables(0).Rows(0).Item(6).ToString = "2" Then
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "4")
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "5")
                            End If
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "A5[0]", "")
                        End If
                    End If
                End If


                '========= SECTION B =========

                dr = datHandler.GetDataReader("Select TC_KEY, TC_STATUTORY_INCOME, (cdbl(TC_DIVIDEND) + cdbl(TC_BUSINESSLOSS_BF)), TC_AGGREGATE_BUS_INCOME," _
                                & " TC_EMPLOYMENT_INCOME, TC_DIVIDEND, (cdbl(TC_INTEREST) + cdbl(TC_DISCOUNT)), " _
                                & " (cdbl(TC_RENTAL_ROYALTY)+cdbl(TC_PREMIUM)), TC_PENSION_AND_ETC," _
                                & " (cdbl(TC_OTHER_GAIN_PROFIT) + cdbl(TC_SEC4A)), TC_ADDITION_43," _
                                & " TC_AGGREGATE_OTHER_SRC, TC_AGGREGATE_INCOME, TC_BUSINESSLOSS_CY," _
                                & " TC_TOTAL1, TC_4, TC_3, TC_TOTAL_INCOME_2, TC_INCOME_TRANSFER_FROM_HW, TC_CHARGEABLE_INCOME," _
                                & " TC_TAX_FIRST_INCOME, TC_TAX_FIRST_TAX, TC_TAX_BALANCE_INCOME, TC_TAX_BALANCE_RATE, TC_TAX_BALANCE_TAX," _
                                & " TC_TAX_SCH1_TAX, TC_TOTAL_INCOME_TAX," _
                                & " (cdbl(TC_INTEREST)+cdbl(TC_DISCOUNT)+cdbl(TC_RENTAL_ROYALTY)+cdbl(TC_PREMIUM)+cdbl(TC_PENSION_AND_ETC)+cdbl(TC_OTHER_GAIN_PROFIT)+cdbl(TC_SEC4A))," _
                                & " TC_REBATES, TC_DONATION_GIFT, TC_RELIEF," _
                                & " (cdbl(TC_INCOME_TAX_CHARGED) - cdbl(TC_TAX_SCH1_TAX)) as TC_TAX_CHARGED," _
                                & " TC_SEC110_DIVIDEND, TC_SEC110_OTHERS, (cdbl(TC_1)+cdbl(TC_2)) as TC_132_133, TC_5," _
                                & " TC_TAX_PAYABLE, TC_TAX_REPAYMENT, (cdbl(TC_INSTALLMENT_PAYMENT_SELF)+cdbl(TC_INSTALLMENT_PAYMENT_HW)) as B18," _
                                & " TC_BALANCE_TAX_PAYABLE, TC_BALANCE_TAX_OVERPAID" _
                                & " from tax_computation where" _
                                & " tc_ref_no ='" & pdfForm.GetRefNo & "' and tc_ya ='" & pdfForm.GetYA & "'")

                If dr.Read() Then

                    '--- B1 ---
                    If Not IsDBNull(dr("TC_EMPLOYMENT_INCOME")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B1[0]", FormatFixedAmount(dr("TC_EMPLOYMENT_INCOME")))
                    End If

                    '--- B2 ---
                    dr2 = datHandler.GetDataReader("Select OS_DV_STAT_INCOME, OS_RT_RENTAL_BF" _
                    & " from income_othersource where" _
                    & " os_ref_no ='" & pdfForm.GetRefNo & "' and os_ya ='" & pdfForm.GetYA & "'")
                    If dr2.Read() Then
                        If Not (IsDBNull(dr2.Item(0)) And IsDBNull(dr2.Item(1))) Then
                            pdfFormFields.SetField(pdfFieldPath & "B2[0]", Fix((CDbl(dr2.Item(0))) + CDbl(dr2.Item(1))))
                        End If

                        '--- B3 ---
                        If Not IsDBNull(dr.Item(27)) Then
                            pdfFormFields.SetField(pdfFieldPath & "B3[0]", (CDbl(dr.Item(27)) - CDbl(dr2.Item(1))))
                        End If
                    End If
                    dr2.Close()

                    '--- B4 ---
                    If Not IsDBNull(dr("TC_AGGREGATE_INCOME")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B4[0]", FormatFixedAmount(dr("TC_AGGREGATE_INCOME")))
                    End If

                    '--- B5 ---
                    If Not IsDBNull(dr("TC_DONATION_GIFT")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B5[0]", FormatFixedAmount(dr("TC_DONATION_GIFT")))
                    End If

                    '--- B6 ---
                    If Not IsDBNull(dr("TC_3")) And Not IsDBNull(dr("TC_4")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B6[0]", FormatFloatingAmount(CDbl(dr("TC_3")) + CDbl(dr("TC_4")), False))
                    ElseIf Not IsDBNull(dr("TC_4")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B6[0]", dr("TC_4"))
                    ElseIf Not IsDBNull(dr("TC_3")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B6[0]", dr("TC_3"))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "B6[0]", "0")
                    End If

                    '--- B7 ---
                    If Not IsDBNull(dr("TC_INCOME_TRANSFER_FROM_HW")) Then
                        If Not String.IsNullOrEmpty(dr("TC_INCOME_TRANSFER_FROM_HW").ToString) Then
                            pdfFormFields.SetField(pdfFieldPath & "B7_1[0]", FormatFixedAmount(dr("TC_INCOME_TRANSFER_FROM_HW")))
                        End If
                    End If

                    Dim boolWithBusiness As Boolean = False
                    dr2 = datHandler.GetDataReader("select TP_HW_TYPEOFINCOME" _
                                                   & " from taxp_profile_hw_others where" _
                                                   & " tp_ref_no ='" & pdfForm.GetRefNo & "'")
                    While dr2.Read()
                        If dr2("TP_HW_TYPEOFINCOME") = "1" Then
                            boolWithBusiness = True
                            Exit While
                        End If
                    End While
                    dr2.Close()

                    dr2 = datHandler.GetDataReader("select TP_TYPE_ASSESSMENT, TP_GENDER, TP_ASSESSMENTON, TP_STATUS, TP_HW_TYPEOFINCOME" _
                                   & " from taxp_profile where" _
                                   & " (tp_ref_no1 + tp_ref_no2 + tp_ref_no3) ='" & pdfForm.GetRefNo & "'")
                    If dr2.Read() Then
                        If Not IsDBNull(dr2("TP_HW_TYPEOFINCOME")) Then
                            If Not String.IsNullOrEmpty(dr2("TP_HW_TYPEOFINCOME").ToString) Then
                                If dr2("TP_TYPE_ASSESSMENT") = "1" Then
                                    If (dr2("TP_GENDER") = "1" And dr2("TP_ASSESSMENTON") = "1") Or _
                                        (dr2("TP_GENDER") = "2" And dr2("TP_ASSESSMENTON") = "2") Then
                                        If dr2("TP_HW_TYPEOFINCOME").ToString = "1" Or boolWithBusiness = True Then
                                            pdfFormFields.SetField(pdfFieldPath & "B7_2[0]", "1")
                                        Else
                                            pdfFormFields.SetField(pdfFieldPath & "B7_2[0]", "2")
                                        End If
                                    Else
                                        pdfFormFields.SetField(pdfFieldPath & "B7_2[0]", "")
                                    End If
                                Else
                                    pdfFormFields.SetField(pdfFieldPath & "B7_2[0]", "")
                                End If
                            End If
                        End If

                        '--- B8 ---
                        If Not IsDBNull(dr2("TP_STATUS")) Then
                            If dr2("TP_STATUS") = "1" Then
                                pdfFormFields.SetField(pdfFieldPath & "B8[0]", "0")
                            ElseIf dr2("TP_STATUS") = "2" And CDbl(pdfForm.GetYA) >= 2007 And dr2("TP_TYPE_ASSESSMENT") = "3" Then
                                pdfFormFields.SetField(pdfFieldPath & "B8[0]", "0")
                            ElseIf dr2("TP_STATUS") = "2" And CDbl(pdfForm.GetYA) = 2006 And dr2("TP_TYPE_ASSESSMENT") = "1" And CDbl(dr("TC_INCOME_TRANSFER_FROM_HW")) = 0 Then
                                pdfFormFields.SetField(pdfFieldPath & "B8[0]", "0")
                            ElseIf (dr2("TP_GENDER") = "1" And dr2("TP_STATUS") = "2" And dr2("TP_TYPE_ASSESSMENT") = "1" And _
                                dr2("TP_ASSESSMENTON") = "1") Or (dr2("TP_GENDER") = "2" And dr2("TP_STATUS") = "2" And dr2("TP_TYPE_ASSESSMENT") = "1" And dr2("TP_ASSESSMENTON") = "2") Then
                                pdfFormFields.SetField(pdfFieldPath & "B8[0]", FormatFixedAmount((CDbl(dr("TC_TOTAL_INCOME_2")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))).ToString))
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "B8[0]", "0")
                            End If
                        End If
                    End If
                    dr2.Close()

                    '--- B9 ---
                    If Not IsDBNull(dr("TC_RELIEF")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B9[0]", FormatFixedAmount(dr("TC_RELIEF").ToString))
                    End If

                    '--- B10 ---
                    If Not IsDBNull(dr("TC_CHARGEABLE_INCOME")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B10[0]", FormatFixedAmount(dr("TC_CHARGEABLE_INCOME").ToString))
                    End If

                    If Not IsDBNull(dr("TC_AGGREGATE_INCOME")) Then
                        If CDbl(dr("TC_AGGREGATE_INCOME")) <= 96000 Then
                            If CDbl(dr.Item("TC_CHARGEABLE_INCOME") - 2000) <= 0 Then
                                pdfFormFields.SetField(pdfFieldPath & "B10a[0]", "2000")
                                pdfFormFields.SetField(pdfFieldPath & "B10b[0]", "0")
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "B10a[0]", "2000")
                                pdfFormFields.SetField(pdfFieldPath & "B10b[0]", CDbl(dr("TC_CHARGEABLE_INCOME")) - 2000)
                            End If
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "B10a[0]", "0")
                            pdfFormFields.SetField(pdfFieldPath & "B10b[0]", FormatFixedAmount(dr("TC_CHARGEABLE_INCOME").ToString))
                        End If
                    End If

                    '--- B11 ---
                    If Not IsDBNull(dr("TC_TAX_FIRST_INCOME")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B11a_1[0]", FormatFixedAmount(dr("TC_TAX_FIRST_INCOME").ToString))
                    End If

                    If Not IsDBNull(dr("TC_TAX_FIRST_TAX")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B11a_2[0]", FormatFixedAmount(Left(CStr(dr("TC_TAX_FIRST_TAX")), (Len(CStr(dr("TC_TAX_FIRST_TAX"))) - 3))))
                    End If

                    If Not IsDBNull(dr("TC_TAX_FIRST_TAX")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B11a_3[0]", Right(CStr(dr("TC_TAX_FIRST_TAX")), 2))
                    End If

                    If Not IsDBNull(dr("TC_TAX_BALANCE_INCOME")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B11b_1[0]", FormatFixedAmount(dr("TC_TAX_BALANCE_INCOME").ToString))
                    End If

                    If Not IsDBNull(dr("TC_TAX_BALANCE_RATE")) Then
                        If Len(dr("TC_TAX_BALANCE_RATE")) > 1 Then
                            pdfFormFields.SetField(pdfFieldPath & "B11b_2[0]", Mid(CStr(dr("TC_TAX_BALANCE_RATE")), 1, 1) & Space(6) & Mid(CStr(dr("TC_TAX_BALANCE_RATE")), 2, 1))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "B11b_2[0]", Space(9) & dr("TC_TAX_BALANCE_RATE"))
                        End If
                    End If

                    If Not IsDBNull(dr("TC_TAX_BALANCE_TAX")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B11b_3[0]", FormatFixedAmount(Left(CStr(dr("TC_TAX_BALANCE_TAX")), (Len(CStr(dr("TC_TAX_BALANCE_TAX"))) - 3))))
                    End If
                    If Not IsDBNull(dr("TC_TAX_BALANCE_TAX")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B11b_4[0]", Right(CStr(dr("TC_TAX_BALANCE_TAX")), 2))
                    End If

                    '--- B12 ---
                    If Not IsDBNull(dr("TC_TAX_SCH1_TAX")) And dr("TC_TOTAL_INCOME_TAX") > 0 Then
                        'pdfFormFields.SetField(pdfFieldPath & "B12_1[0]", Left(CStr((dr("TC_TOTAL_INCOME_TAX") - dr("TC_TAX_SCH1_TAX"))), (Len(CStr((dr("TC_TOTAL_INCOME_TAX") - dr("TC_TAX_SCH1_TAX")))) - 3)))
                        'pdfFormFields.SetField(pdfFieldPath & "B12_2[0]", Right(CStr(dr("TC_TOTAL_INCOME_TAX") - dr("TC_TAX_SCH1_TAX")), 2))
                        'pdfFormFields.SetField(pdfFieldPath & "B12_1[0]", Left((CDbl(dr("TC_TOTAL_INCOME_TAX")).ToString("0.00") - CDbl(dr("TC_TAX_SCH1_TAX")).ToString("0.00")), (Len(CDbl(dr("TC_TOTAL_INCOME_TAX")).ToString("0.00") - CDbl(dr("TC_TAX_SCH1_TAX")).ToString("0.00")) - 3)))
                        'dannylee 2012Su2.1
                        pdfFormFields.SetField(pdfFieldPath & "B12_1[0]", Left((CDbl(dr("TC_TOTAL_INCOME_TAX")) - CDbl(dr("TC_TAX_SCH1_TAX"))).ToString("0.00"), (Len((CDbl(dr("TC_TOTAL_INCOME_TAX")) - CDbl(dr("TC_TAX_SCH1_TAX"))).ToString("0.00")) - 3)))
                        'end
                        pdfFormFields.SetField(pdfFieldPath & "B12_2[0]", Right(CDbl(dr("TC_TOTAL_INCOME_TAX")).ToString("0.00"), 2))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "B12_1[0]", Left((CDbl(dr("TC_TOTAL_INCOME_TAX")).ToString("0.00")), (Len(CDbl(dr("TC_TOTAL_INCOME_TAX")).ToString("0.00")) - 3)))
                        pdfFormFields.SetField(pdfFieldPath & "B12_2[0]", Right(CDbl(dr("TC_TOTAL_INCOME_TAX")).ToString("0.00"), 2))
                    End If

                    '--- B13 ---
                    dr2 = datHandler.GetDataReader("SELECT TCR_KEY, TCR_AMOUNT FROM [TAX_REBATE] WHERE [TC_KEY]= " & dr("TC_KEY"))
                    While dr2.Read()
                        If Not IsDBNull(dr2("TCR_KEY")) Then
                            If Not String.IsNullOrEmpty(dr2("TCR_KEY").ToString) Then
                                Select Case dr2("TCR_KEY")
                                    Case 1
                                        pdfFormFields.SetField(pdfFieldPath & "B13_1[0]", dr2("TCR_AMOUNT").ToString)
                                    Case 2
                                        pdfFormFields.SetField(pdfFieldPath & "B13_2[0]", dr2("TCR_AMOUNT").ToString)
                                    Case 3
                                        pdfFormFields.SetField(pdfFieldPath & "B13_3[0]", Left(CStr(FormatFloatingAmount((dr2("TCR_AMOUNT")), True)), (Len(CStr(FormatFloatingAmount(dr2("TCR_AMOUNT"), True))) - 2)))
                                        pdfFormFields.SetField(pdfFieldPath & "B13_4[0]", Right(CStr(FormatFloatingAmount((dr2("TCR_AMOUNT")), True)), 2))
                                End Select
                            End If
                        End If
                    End While

                    dr2.Close()

                    If Not IsDBNull(dr("TC_REBATES")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B13_5[0]", Left(CStr(FormatFloatingAmount((dr("TC_REBATES")), True)), (Len(CStr(FormatFloatingAmount(dr("TC_REBATES"), True))) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B13_6[0]", Right(CStr(FormatFloatingAmount((dr("TC_REBATES")), True)), 2))
                    End If

                    '--- B14 ---
                    If Not IsDBNull(dr("TC_TAX_CHARGED")) Then
                        'pdfFormFields.SetField(pdfFieldPath & "B14_1[0]", Left(CStr(FormatFloatingAmount((dr("TC_TAX_CHARGED")), True)), (Len(CStr(FormatFloatingAmount(dr("TC_TAX_CHARGED"), True))) - 2)))
                        'pdfFormFields.SetField(pdfFieldPath & "B14_2[0]", Right(CStr(FormatFloatingAmount((dr("TC_TAX_CHARGED")), True)), 2))
                        pdfFormFields.SetField(pdfFieldPath & "B14_1[0]", Left((CDbl(dr("TC_TAX_CHARGED")).ToString("0.00")), (Len(CDbl(dr("TC_TAX_CHARGED")).ToString("0.00")) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B14_2[0]", Right(CStr(FormatNumber(dr("TC_TAX_CHARGED"), 2)), 2))
                    End If

                    '--- B15 ---
                    If Not IsDBNull(dr("TC_SEC110_DIVIDEND")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B15_1[0]", Left(CStr(FormatFloatingAmount((dr("TC_SEC110_DIVIDEND")), True)), (Len(CStr(FormatFloatingAmount(dr("TC_SEC110_DIVIDEND"), True))) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B15_2[0]", Right(CStr(FormatFloatingAmount((dr("TC_SEC110_DIVIDEND")), True)), 2))
                    End If
                    If Not IsDBNull(dr("TC_SEC110_OTHERS")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B15_3[0]", Left(CStr(FormatFloatingAmount((dr("TC_SEC110_OTHERS")), True)), (Len(CStr(FormatFloatingAmount(dr("TC_SEC110_OTHERS"), True))) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B15_4[0]", Right(CStr(FormatFloatingAmount((dr("TC_SEC110_OTHERS")), True)), 2))
                    End If
                    If Not IsDBNull(dr("TC_132_133")) Then
                        'pdfFormFields.SetField(pdfFieldPath & "B15_5[0]", Left(CStr(FormatFloatingAmount((dr("TC_132_133")), True)), (Len(CStr(FormatFloatingAmount(dr("TC_132_133"), True))) - 2)))
                        'pdfFormFields.SetField(pdfFieldPath & "B15_6[0]", Right(CStr(FormatFloatingAmount((dr("TC_132_133")), True)), 2))
                        pdfFormFields.SetField(pdfFieldPath & "B15_5[0]", Left((CDbl(dr("TC_132_133")).ToString("0.00")), (Len(CDbl(dr("TC_132_133")).ToString("0.00")) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B15_6[0]", Right(CStr(FormatNumber(dr("TC_132_133"), 2)), 2))
                    End If
                    If Not IsDBNull(dr("TC_5")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B15_7[0]", Left(CStr(FormatFloatingAmount((dr("TC_5")), True)), (Len(CStr(FormatFloatingAmount(dr("TC_5"), True))) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B15_8[0]", Right(CStr(FormatFloatingAmount((dr("TC_5")), True)), 2))
                    End If

                    '--- B16 ---
                    If Not IsDBNull(dr("TC_TAX_PAYABLE")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B16_1[0]", Left(CStr(FormatFloatingAmount((dr("TC_TAX_PAYABLE")), True)), (Len(CStr(FormatFloatingAmount(dr("TC_TAX_PAYABLE"), True))) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B16_2[0]", Right(CStr(FormatFloatingAmount((dr("TC_TAX_PAYABLE")), True)), 2))
                    End If

                    '--- B17 ---
                    If Not IsDBNull(dr("TC_TAX_REPAYMENT")) Then
                        pdfFormFields.SetField(pdfFieldPath & "B17_1[0]", Left(CStr(FormatFloatingAmount((dr("TC_TAX_REPAYMENT")), True)), (Len(CStr(FormatFloatingAmount(dr("TC_TAX_REPAYMENT"), True))) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B17_2[0]", Right(CStr(FormatFloatingAmount((dr("TC_TAX_REPAYMENT")), True)), 2))
                    End If

                    '--- B18 ---
                    If Not IsDBNull(dr("B18")) Then
                        'pdfFormFields.SetField(pdfFieldPath & "B18_1[0]", Left(CStr(FormatNumber(CDbl(dr("B18")), 2)), (Len(CStr(FormatNumber(CDbl(dr("B18")), 2)) - 2))))
                        pdfFormFields.SetField(pdfFieldPath & "B18_1[0]", Left((CDbl(dr("B18")).ToString("0.00")), (Len(CDbl(dr("B18")).ToString("0.00")) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B18_2[0]", Right(CStr(FormatNumber(dr("B18"), 2)), 2))
                    End If

                    '--- B19 ---
                    If CDbl(dr("TC_TAX_PAYABLE")) >= CDbl(dr("B18")) Then
                        If Not IsDBNull(dr("TC_BALANCE_TAX_PAYABLE")) Then
                            If Not String.IsNullOrEmpty(dr("TC_BALANCE_TAX_PAYABLE").ToString) Then
                                pdfFormFields.SetField(pdfFieldPath & "B19_1[0]", "")
                                pdfFormFields.SetField(pdfFieldPath & "B19_2[0]", Left(CStr(FormatFloatingAmount((dr("TC_BALANCE_TAX_PAYABLE")), True)), (Len(CStr(FormatFloatingAmount(dr("TC_BALANCE_TAX_PAYABLE"), True))) - 2)))
                                pdfFormFields.SetField(pdfFieldPath & "B19_3[0]", Right(CStr(FormatFloatingAmount((dr("TC_BALANCE_TAX_PAYABLE")), True)), 2))
                            End If
                        End If
                    Else
                        If Not IsDBNull(dr("TC_BALANCE_TAX_OVERPAID")) Then
                            If Not String.IsNullOrEmpty(dr("TC_BALANCE_TAX_OVERPAID").ToString) Then
                                pdfFormFields.SetField(pdfFieldPath & "B19_1[0]", "X")
                                pdfFormFields.SetField(pdfFieldPath & "B19_2[0]", Left(CStr(FormatFloatingAmount((dr("TC_BALANCE_TAX_OVERPAID")), True)), (Len(CStr(FormatFloatingAmount(dr("TC_BALANCE_TAX_OVERPAID"), True))) - 2)))
                                pdfFormFields.SetField(pdfFieldPath & "B19_3[0]", Right(CStr(FormatFloatingAmount((dr("TC_BALANCE_TAX_OVERPAID")), True)), 2))
                            End If
                        End If
                    End If
                End If

                dr.Close()

                dr = datHandler.GetDataReader("select tp_last_passport_no, TP_WORKER_APPROVEDATE, TP_COM_ADD_STATUS from taxp_profile2 where" _
                        & " tp_ref_no= '" & pdfForm.GetRefNo & "'")


                '========= DECLARATION =========

                If Not String.IsNullOrEmpty(pdfForm.GetDeclarationReturn) Then
                    pdfFormFields.SetField(pdfFieldPath & "AK3", pdfForm.GetDeclarationReturn)

                    If pdfForm.GetDeclarationReturn = "1" Then
                        If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) Then
                            If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(0).ToString()) Then
                                pdfFormFields.SetField(pdfFieldPath & "AK1[0]", ds.Tables(0).Rows(0).Item(0).ToString.ToUpper)
                            End If
                        End If

                        If Not IsDBNull(ds.Tables(0).Rows(0).Item(2)) Then
                            If Not String.IsNullOrEmpty(Trim(ds.Tables(0).Rows(0).Item(2).ToString)) Then
                                pdfFormFields.SetField(pdfFieldPath & "AK2[0]", FormatICNumber(ds.Tables(0).Rows(0).Item(2).ToString))
                            Else
                                If dr.Read() Then
                                    If Not IsDBNull(dr("tp_last_passport_no")) Then
                                        If Not String.IsNullOrEmpty(dr("tp_last_passport_no").ToString) Then
                                            pdfFormFields.SetField(pdfFieldPath & "AK2[0]", dr("tp_last_passport_no").ToString)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        dr2 = datHandler.GetDataReader("SELECT TP_STATUS FROM TAXP_PROFILE WHERE (TP_REF_NO1 + TP_REF_NO2 + TP_REF_NO3)='" & pdfForm.GetRefNo & "'")
                        If dr2.Read() Then
                            If dr2("TP_STATUS") = 4 Then
                                pdfFormFields.SetField(pdfFieldPath & "AK3", "3")
                            End If
                        End If
                        dr2.Close()
                        If Not String.IsNullOrEmpty(pdfForm.GetDeclarationBy.ToString) Then
                            pdfFormFields.SetField(pdfFieldPath & "AK1", pdfForm.GetDeclarationBy)
                        End If
                        If Not String.IsNullOrEmpty(pdfForm.GetDeclarationID.ToString) Then
                            pdfFormFields.SetField(pdfFieldPath & "AK2", pdfForm.GetDeclarationID)
                        End If
                    End If
                End If

                If Not String.IsNullOrEmpty(pdfForm.GetDeclarationDate) Then
                    pdfFormFields.SetField(pdfFieldPath & "AK4[0]", FormatDeclarationDate(pdfForm.GetDeclarationDate))
                End If



                '========= SECTION C =========

                '--- C1 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(10)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(10).ToString()) Then
                        pdfFormFields.SetField(pdfFieldPath & "C1[0]", ds.Tables(0).Rows(0).Item(10).ToString.ToUpper)
                    End If
                End If

                '--- C2 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(11)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(11).ToString()) Then
                        pdfFormFields.SetField(pdfFieldPath & "C2[0]", FormatICNumber(ds.Tables(0).Rows(0).Item(11).ToString))
                    End If
                End If

                '--- C3 ---
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(12)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(12).ToString()) Then
                        pdfFormFields.SetField(pdfFieldPath & "C3[0]", ds.Tables(0).Rows(0).Item(12).ToString)
                    End If
                End If
            End If
            dr.Close()
            ds.Clear()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub Page2()

        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As OleDbDataReader = Nothing
        Dim dr2 As OleDbDataReader = Nothing
        Dim prmOledb(0) As OleDbParameter
        Dim strHWIC As String = ""
        Dim strArray(1) As String
        Dim intCounter As Integer = 1
        Dim dblTotalValue As Double = 0
        Dim dblTotalValue1 As Double = 0
        Dim intArrayChild50(5) As Integer
        Dim intArrayChild100(5) As Integer


        Try
            prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page2[0]."

            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Name2", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ref2", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()


            ' ==== Master Data ==== "
            ds = datHandler.GetData("select tp_tel1, tp_tel2, tp_mobile1, tp_mobile2," _
            & " (tp_employer_no2 + tp_employer_no3), tp_bank, tp_bank_acc, tp_email" _
            & " from taxp_profile where (tp_ref_no1 + tp_ref_no2 + tp_ref_no3)=?", prmOledb)


            '========= SECTION D =========

            '--- D1 ---
            If ds.Tables(0).Rows.Count > 0 Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) And Not IsDBNull(ds.Tables(0).Rows(0).Item(1)) Then
                    pdfFormFields.SetField(pdfFieldPath & "D1[0]", ds.Tables(0).Rows(0).Item(0) & "-" & ds.Tables(0).Rows(0).Item(1))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "D1[0]", ds.Tables(0).Rows(0).Item(2) & "-" & ds.Tables(0).Rows(0).Item(3))
                End If
            End If

            '--- D2 ---
            If ds.Tables(0).Rows.Count > 0 Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(4)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(4).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "D2[0]", ds.Tables(0).Rows(0).Item(4).ToString)
                    End If
                End If
            End If

            '--- D3 ---
            If ds.Tables(0).Rows.Count > 0 Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(5)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(5).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "D3[0]", ds.Tables(0).Rows(0).Item(5).ToString.ToUpper)
                    End If
                End If
            End If

            '--- D4 ---
            If ds.Tables(0).Rows.Count > 0 Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(6)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(6).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "D4[0]", ds.Tables(0).Rows(0).Item(6).ToString)
                    End If
                End If
            End If

            '--- D5 ---
            If ds.Tables(0).Rows.Count > 0 Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(7)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(7).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "D5[0]", ds.Tables(0).Rows(0).Item(7).ToString)
                    End If
                End If
            End If


            '========= SECTION E ==========

            '--- E1 & E2 ---
            intCounter = 1
            pdfFormFields.SetField(pdfFieldPath & "E1_1[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "E2_1[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "E1_2[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "E2_2[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "E1_3[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "E2_3[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "E1_4[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "E2_4[0]", "0")

            dr = datHandler.GetDataReader("Select * from preceding_year where py_ref_no= '" & pdfForm.GetRefNo & "' and py_ya= '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                dr2 = datHandler.GetDataReader("Select TOP 2 PY_INCOME_TYPE, PY_PAYMENT_YEAR, PY_AMOUNT, PY_EPF" _
                                                    & " From PRECEDING_YEAR_DETAIL Where" _
                                                    & " PY_KEY= " & dr("PY_KEY") & " Order By PY_DKEY")
                While dr2.Read()

                    If Not IsDBNull(dr2("PY_INCOME_TYPE")) Then
                        pdfFormFields.SetField(pdfFieldPath & "E" & intCounter.ToString & "_1[0]", dr2("PY_INCOME_TYPE").ToString.ToUpper)
                    End If
                    If Not IsDBNull(dr2("PY_PAYMENT_YEAR")) Then
                        pdfFormFields.SetField(pdfFieldPath & "E" & intCounter.ToString & "_2[0]", dr2("PY_PAYMENT_YEAR").ToString)
                    End If
                    If Not IsDBNull(dr2("PY_AMOUNT")) Then
                        pdfFormFields.SetField(pdfFieldPath & "E" & intCounter.ToString & "_3[0]", FormatFixedAmount(dr2("PY_AMOUNT").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "E" & intCounter.ToString & "_3[0]", "0")
                    End If
                    If Not IsDBNull(dr2("PY_EPF")) Then
                        pdfFormFields.SetField(pdfFieldPath & "E" & intCounter.ToString & "_4[0]", FormatFixedAmount(dr2("PY_EPF").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "E" & intCounter.ToString & "_4[0]", "0")

                    End If
                    intCounter = intCounter + 1
                End While
                dr2.Close()
            End If
            dr.Close()


            '========= SECTION F =========

            dr = datHandler.GetDataReader("Select TC_KEY, TC_STATUTORY_INCOME, TC_BUSINESSLOSS_BF, TC_AGGREGATE_BUS_INCOME," _
                    & " TC_EMPLOYMENT_INCOME, TC_DIVIDEND, (cdbl(TC_INTEREST) + cdbl(TC_DISCOUNT)), " _
                    & " (cdbl(TC_RENTAL_ROYALTY)+cdbl(TC_PREMIUM)), TC_PENSION_AND_ETC," _
                    & " (cdbl(TC_OTHER_GAIN_PROFIT) + cdbl(TC_SEC4A)), TC_ADDITION_43," _
                    & " TC_AGGREGATE_OTHER_SRC, TC_AGGREGATE_INCOME, TC_BUSINESSLOSS_CY," _
                    & " TC_TOTAL1, TC_4, TC_3, TC_TOTAL_INCOME_2, TC_INCOME_TRANSFER_FROM_HW, TC_RELIEF" _
                    & " from tax_computation where" _
                    & " tc_ref_no ='" & pdfForm.GetRefNo & "' and tc_ya ='" & pdfForm.GetYA & "'")

            If dr.Read() Then


                dblTotalValue = 0
                dr2 = datHandler.GetDataReader("select tcc_key, tcc_amount from tax_relief where " _
                                   & " tc_key =" & dr("TC_KEY") & " order by tcc_key")

                While dr2.Read()
                    If Not IsDBNull(dr2("tcc_key")) And Not IsDBNull(dr2("tcc_amount")) Then
                        If Not String.IsNullOrEmpty(dr2("tcc_amount")) Then

                            Select Case dr2("tcc_key")
                                Case 2
                                    pdfFormFields.SetField(pdfFieldPath & "F2[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 3
                                    pdfFormFields.SetField(pdfFieldPath & "F3[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 4
                                    pdfFormFields.SetField(pdfFieldPath & "F4[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 5
                                    pdfFormFields.SetField(pdfFieldPath & "F5[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 6
                                    pdfFormFields.SetField(pdfFieldPath & "F6[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                    dblTotalValue = dblTotalValue + CDbl(dr2("tcc_amount"))
                                Case 7
                                    pdfFormFields.SetField(pdfFieldPath & "F7[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                    dblTotalValue = dblTotalValue + CDbl(dr2("tcc_amount"))
                                Case 8
                                    pdfFormFields.SetField(pdfFieldPath & "F8[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 21
                                    pdfFormFields.SetField(pdfFieldPath & "F9[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 22
                                    pdfFormFields.SetField(pdfFieldPath & "F10[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 23
                                    pdfFormFields.SetField(pdfFieldPath & "F11[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 24
                                    pdfFormFields.SetField(pdfFieldPath & "F12[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 9
                                    pdfFormFields.SetField(pdfFieldPath & "F13[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 10
                                    pdfFormFields.SetField(pdfFieldPath & "F14[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 13
                                    pdfFormFields.SetField(pdfFieldPath & "F15[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 14
                                    pdfFormFields.SetField(pdfFieldPath & "F16a_5[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 15
                                    pdfFormFields.SetField(pdfFieldPath & "F16b_9[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 16
                                    pdfFormFields.SetField(pdfFieldPath & "F16c_9[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 17
                                    pdfFormFields.SetField(pdfFieldPath & "F17[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 25
                                    pdfFormFields.SetField(pdfFieldPath & "F18[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                    dblTotalValue1 = dblTotalValue1 + CDbl(dr2("tcc_amount"))
                                Case 18
                                    pdfFormFields.SetField(pdfFieldPath & "F19[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                            End Select
                        End If
                    End If
                End While

                dr2.Close()

                If Not IsDBNull(dr("TC_RELIEF")) Then
                    pdfFormFields.SetField(pdfFieldPath & "F20[0]", FormatFixedAmount(dr("TC_RELIEF").ToString))
                End If

            End If
            dr.Close()

            dr = datHandler.GetDataReader("Select TC_KEY, TC_RELIEF, TC_CHARGEABLE_INCOME, TC_TAX_FIRST_INCOME, TC_TAX_FIRST_TAX," _
                                & " TC_TAX_BALANCE_INCOME, TC_TAX_BALANCE_RATE, TC_TAX_BALANCE_TAX, TC_TOTAL_INCOME_TAX," _
                                & " TC_REBATES, TC_INCOME_TAX_CHARGED, TC_SEC110_DIVIDEND,TC_SEC110_OTHERS, TC_1," _
                                & " TC_2, TC_TAX_PAYABLE, TC_TAX_REPAYMENT, TC_TAX_SCH1_INCOME, TC_TAX_SCH1_TAX" _
                                & " From TAX_COMPUTATION Where" _
                                & " TC_REF_NO= '" & pdfForm.GetRefNo & "' and TC_YA= '" & pdfForm.GetYA & "'")
            If dr.Read() Then


                dblTotalValue = 0


                ReDim intArrayChild50(5)
                ReDim intArrayChild100(5)

                'For i As Integer = 1 To 2
                '    intCounter(i) = 0
                'Next
                'For i As Integer = 1 To 5
                '    intArrayChild50(i) = 0
                '    intArrayChild100(i) = 0
                'Next

                dr2 = datHandler.GetDataReader("SELECT TCC_KEY, TCC_100, TCC_50 FROM [TAX_RELIEF_CHILD] WHERE [TC_KEY] = " & dr("TC_KEY") & " order by  [TCC_KEY]")
                While dr2.Read()
                    If Not IsDBNull(dr2("TCC_KEY")) Then
                        Select Case dr2("TCC_KEY")
                            Case 14
                                If Not IsDBNull(dr2("TCC_100")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_100"))) And Not Trim(dr2("TCC_100")) = "" Then
                                    If CDbl(Trim(dr2("TCC_100"))) = 1000 Then
                                        intArrayChild100(1) = intArrayChild100(1) + 1
                                    End If
                                End If
                                If Not IsDBNull(dr2("TCC_50")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_50"))) And Not Trim(dr2("TCC_50")) = "" Then
                                    If CDbl(Trim(dr2("TCC_50"))) = 500 Then
                                        intArrayChild50(1) = intArrayChild50(1) + 1
                                    End If
                                End If
                            Case 15
                                If Not IsDBNull(dr2("TCC_100")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_100"))) And Not Trim(dr2("TCC_100")) = "" Then
                                    If CDbl(Trim(dr2("TCC_100"))) = 1000 Then
                                        intArrayChild100(2) = intArrayChild100(2) + 1
                                    ElseIf CDbl(Trim(dr2("TCC_100"))) = 6000 Then
                                        intArrayChild100(3) = intArrayChild100(3) + 1
                                    End If
                                End If
                                If Not IsDBNull(dr2("TCC_50")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_50"))) And Not Trim(dr2("TCC_50")) = "" Then
                                    If CDbl(Trim(dr2("TCC_50"))) = 500 Then
                                        intArrayChild50(2) = intArrayChild50(2) + 1
                                    ElseIf CDbl(Trim(dr2("TCC_50"))) = 3000 Then
                                        intArrayChild50(3) = intArrayChild50(3) + 1
                                    End If
                                End If
                            Case 16
                                If Not IsDBNull(dr2("TCC_100")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_100"))) And Not Trim(dr2("TCC_100")) = "" Then
                                    If CDbl(Trim(dr2("TCC_100"))) = 5000 Then
                                        intArrayChild100(4) = intArrayChild100(4) + 1
                                    ElseIf CDbl(Trim(dr2("TCC_100"))) = 11000 Then
                                        intArrayChild100(5) = intArrayChild100(5) + 1
                                    End If
                                End If
                                If Not IsDBNull(dr2("TCC_50")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_50"))) And Not Trim(dr2("TCC_50")) = "" Then
                                    If CDbl(Trim(dr2("TCC_50"))) = 2500 Then
                                        intArrayChild50(4) = intArrayChild50(4) + 1
                                    ElseIf CDbl(Trim(dr2("TCC_50"))) = 5500 Then
                                        intArrayChild50(5) = intArrayChild50(5) + 1
                                    End If
                                End If
                        End Select
                    End If
                End While
                dr2.Close()

                For i As Integer = 1 To 5
                    Select Case i
                        Case 1
                            pdfFormFields.SetField(pdfFieldPath & "F16a_1[0]", intArrayChild100(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "F16a_2[0]", intArrayChild100(i) * 1000)
                            pdfFormFields.SetField(pdfFieldPath & "F16a_3[0]", intArrayChild50(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "F16a_4[0]", intArrayChild50(i) * 500)
                        Case 2
                            pdfFormFields.SetField(pdfFieldPath & "F16b_1[0]", intArrayChild100(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "F16b_2[0]", intArrayChild100(i) * 1000)
                            pdfFormFields.SetField(pdfFieldPath & "F16b_3[0]", intArrayChild50(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "F16b_4[0]", intArrayChild50(i) * 500)
                        Case 3
                            pdfFormFields.SetField(pdfFieldPath & "F16b_5[0]", intArrayChild100(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "F16b_6[0]", intArrayChild100(i) * 6000)
                            pdfFormFields.SetField(pdfFieldPath & "F16b_7[0]", intArrayChild50(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "F16b_8[0]", intArrayChild50(i) * 3000)
                        Case 4
                            pdfFormFields.SetField(pdfFieldPath & "F16c_1[0]", intArrayChild100(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "F16c_2[0]", intArrayChild100(i) * 5000)
                            pdfFormFields.SetField(pdfFieldPath & "F16c_3[0]", intArrayChild50(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "F16c_4[0]", intArrayChild50(i) * 2500)
                        Case 5
                            pdfFormFields.SetField(pdfFieldPath & "F16c_5[0]", intArrayChild100(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "F16c_6[0]", intArrayChild100(i) * 11000)
                            pdfFormFields.SetField(pdfFieldPath & "F16c_7[0]", intArrayChild50(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "F16c_8[0]", intArrayChild50(i) * 5500)
                    End Select
                    'intCounter(0) = intCounter(0) + intArrayChild50(i) + intArrayChild100(i)
                Next
            End If
            dr.Close()


            '========= SECTION G =========

            ReDim strArray(0)
            dr = datHandler.GetDataReader("SELECT * FROM [TAXA_PROFILE] Where [TA_KEY] =" & pdfForm.GetTaxAgent)
            If dr.Read() Then

                '--- G1 ---
                If Not IsDBNull(dr("TA_CO_NAME")) Then
                    If Not String.IsNullOrEmpty(dr("TA_CO_NAME").ToString) Then
                        strArray = SplitText(dr("TA_CO_NAME").ToString, 40)
                        pdfFormFields.SetField(pdfFieldPath & "G1_1", strArray(0).ToUpper)
                        pdfFormFields.SetField(pdfFieldPath & "G1_2", strArray(1).ToUpper)
                    End If
                End If

                '--- G2 ---
                If Not IsDBNull(dr("TA_TEL_NO")) And Not IsDBNull(dr("TA_MOBILE")) Then
                    pdfFormFields.SetField(pdfFieldPath & "G2", FormatPhoneNumber("", dr("TA_TEL_NO").ToString, "", dr("TA_MOBILE").ToString))
                End If

                '--- G3 ---
                If Not IsDBNull(dr("TA_LICENSE")) Then
                    pdfFormFields.SetField(pdfFieldPath & "G3", dr("TA_LICENSE").ToString)
                End If

            End If
            dr.Close()

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
            Select Case (de.Key.ToString.Remove(0, 18))
                Case "Page6[0].NyataA[0]"
                    CheckFieldEmpty(de.Key.ToString, 29)
                Case "Page3[0].I_1[0]", _
                    "Page3[0].I_2[0]", _
                    "Page3[0].A14[0]", _
                    "Page3[0].A13[0]", _
                    "Page3[0].B1_1[0]", _
                    "Page3[0].B1_2[0]", _
                    "Page6[0].H1_1[0]", _
                    "Page6[0].H1_2[0]", _
                    "Page6[0].Akuan1_1[0]", _
                    "Page6[0].Akuan1_2[0]"
                    CheckFieldEmpty(de.Key.ToString, 28)
                Case "Page3[0].A7_1[0]", _
                    "Page3[0].A7_2[0]", _
                    "Page3[0].A7_3[0]"
                    CheckFieldEmpty(de.Key.ToString, 26)
                Case "Page3[0].A8[0]", _
                    "Page6[0].Nyatab[0]"
                    CheckFieldEmpty(de.Key.ToString, 13)
                Case "Page3[0].A4[0]", _
                    "Page3[0].A7a[0]", _
                    "Page6[0].Akuan4[0]", _
                    "Page4[0].H3[0]", _
                    "Page4[0].H5[0]", _
                    "Page6[0].NyataTarikh[0]"
                    CheckFieldEmpty(de.Key.ToString, 8)
                Case "Page3[0].II_1[0]", _
                    "Page3[0].A1[0]", _
                    "Page3[0].B2_1[0]", _
                    "Page5[0].D11_1[0]", _
                    "Page5[0].D11_2[0]", _
                    "Page5[0].D11_3[0]", _
                    "Page5[0].D11a_1[0]", _
                    "Page5[0].D11a_3[0]", _
                    "Page5[0].D11b1_1[0]", _
                    "Page5[0].D11b1_3[0]", _
                    "Page5[0].D11b2_1[0]", _
                    "Page5[0].D11b2_3[0]", _
                    "Page5[0].D11c1_1[0]", _
                    "Page5[0].D11c1_3[0]", _
                    "Page5[0].D11c2_1[0]", _
                    "Page5[0].D11c2_3[0]", _
                    "Page5[0].E2b_2[0]"
                    pdfFormFields.SetField(de.Key.ToString, RTrim("--"))
                Case "Page3[0].A2[0]", _
                    "Page3[0].A3[0]", _
                    "Page3[0].A5[0]", _
                    "Page3[0].A6[0]", _
                    "Page4[0].C17_1[0]", _
                    "Page6[0].Akuan3[0]"
                    pdfFormFields.SetField(de.Key.ToString, RTrim(""))
                Case Else
                    pdfFormFields.SetField(de.Key.ToString, RTrim("---"))
            End Select


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
            strTemp = Format(dtTemp, "dd-MM-yyyy")
        End If
        Return strTemp

    End Function

    Protected Function FormatDeclarationDate(ByVal strTemp As String) As String

        If Not strTemp = "" Then
            strTemp = strTemp.Insert(2, "-").Insert(5, "-")
        End If
        Return strTemp

    End Function

    ''' <summary>
    ''' Format the Amount String to 0000#
    ''' </summary>
    ''' <param name="strTemp">The specific Amount</param>
    ''' <returns>Modified Amount</returns>
    ''' <remarks></remarks>
    Protected Function FormatFixedAmount(ByVal strTemp As String) As String

        If Not strTemp = "" Then
            If CDbl(strTemp) > 0 Then
                strTemp = Math.Ceiling(CDbl(strTemp)).ToString.Replace(",", "")
            Else
                strTemp = "0"
            End If
        End If
        Return strTemp

    End Function

    ''' <summary>
    ''' Format the Amount String to 000#
    ''' </summary>
    ''' <param name="strTemp">The specific Amount</param>
    ''' <returns>Modified Amount</returns>
    ''' <remarks></remarks>
    Protected Function FormatFloatingAmount(ByVal strTemp As String, ByVal intFloating As Boolean) As String

        If intFloating = True Then
            If Not strTemp = "" Then
                If CDbl(strTemp) > 0 Then
                    strTemp = strTemp.ToString.Replace(",", "").Replace(".", "")
                Else
                    strTemp = "000"
                End If
            End If
        Else
            If Not strTemp = "" Then
                If CDbl(strTemp) > 0 Then
                    strTemp = Math.Ceiling(CDbl(strTemp)).ToString.Replace(",", "")
                Else
                    strTemp = "0"
                End If
            End If
        End If
        Return strTemp

    End Function

    ''' <summary>
    ''' Format the String Number type to specify format
    ''' </summary>
    ''' <param name="strTemp">The specific string number</param>
    ''' <param name="intMaxChar">Max of the specific string number to have.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function FormatStringNumber(ByVal strTemp As String, ByVal intMaxChar As Integer)
        Dim strFormat As String = ""
        If Not String.IsNullOrEmpty(strTemp) Or Not strTemp = "" Then
            For i As Integer = 0 To intMaxChar - 1
                strFormat = Trim(strFormat) & "0"
            Next
            strTemp = Mid(CLng(strTemp).ToString(strFormat), 1, intMaxChar)
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
            strTemp = (strHomePrefix & strHome).Replace("-", "")

        ElseIf Not String.IsNullOrEmpty(strMobile) Or strMobile = " " Then
            If strHomePrefix.Length = 2 Then
                strMobilePrefix = " " & strMobilePrefix
            End If
            strTemp = (strMobilePrefix & strMobile).Replace("-", "")
        End If
        Return strTemp

    End Function

    Protected Function GetStatusOfTax() As String
        Dim ds As New DataSet
        Dim dr As OleDbDataReader = Nothing
        Dim prmOledb(0) As OleDbParameter
        Dim strStatus As String = ""

        ReDim prmOledb(1)
        prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
        prmOledb(1) = New OleDbParameter("@ya", pdfForm.GetYA)
        ds = datHandler.GetData("SELECT TC_BALANCE_TAX_PAYABLE, TC_BALANCE_TAX_OVERPAID, TC_TAX_REPAYMENT FROM TAX_COMPUTATION WHERE TC_REF_NO=? AND TC_YA=?", prmOledb)
        If ds.Tables.Count > 0 Then
            If ds.Tables(0).Rows.Count > 0 Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) And Not IsDBNull(ds.Tables(0).Rows(0).Item(1)) Then
                    If CDbl(ds.Tables(0).Rows(0).Item(1).ToString) > 0 Then
                        strStatus = "EXCESS"
                    ElseIf CDbl(ds.Tables(0).Rows(0).Item(0).ToString) > 0 Then
                        strStatus = "BALANCE"
                    ElseIf CDbl(ds.Tables(0).Rows(0).Item(2).ToString) > 0 Then
                        strStatus = "REPAYABLE"
                    ElseIf CDbl(ds.Tables(0).Rows(0).Item(0).ToString) = 0 And CDbl(ds.Tables(0).Rows(0).Item(1).ToString) = 0 Then
                        strStatus = "NIL"
                    End If
                End If
            End If
        End If
        Return strStatus
        ds.Dispose()
    End Function

    Protected Function FormatICNumber(ByVal strTemp As String) As String

        If Not Trim(strTemp) = "" Then
            strTemp = strTemp.Insert(6, "-").Insert(9, "-")
        End If
        Return strTemp

    End Function

    Protected Function FormatAddress(ByVal strTemp As String) As String

        If Not Trim(strTemp) = "" Then
            strTemp = strTemp.ToString.Replace(", ,", ",").Replace(", ,", ",")
        End If
        Return strTemp

    End Function

#End Region
End Class

