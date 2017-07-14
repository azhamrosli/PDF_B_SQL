Imports iTextSharp.text.pdf
Imports System.Data.OleDb
Imports System.IO

Public Class clsBorangBE2011
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
        Page3()
        Page4()
        Page5()

        pdfForm.OpenFile()
        pdfForm.CloseStamper()
    End Sub

#End Region

#Region "Insert the page function here"

    Private Sub Page1()

        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As OleDbDataReader = Nothing
        Dim prmOledb(0) As OleDbParameter
        Dim strHWIC As String = ""
        Dim strArray(1) As String


        Try
            prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page3[0]."


            ' ==== Master Data ==== "
            ds = datHandler.GetData("select tp_name , tp_ref_no_prefix, (tp_ref_no1 + tp_ref_no2 + tp_ref_no3)," _
                        & " (tp_ic_new_1 + tp_ic_new_2 + tp_ic_new_3), tp_ic_old," _
                        & " tp_police_no, tp_army_no, tp_passport_no," _
                        & " tp_country, tp_gender, tp_status, tp_date_marriage," _
                        & " tp_date_divorce, tp_type_assessment, tp_kup," _
                        & " (tp_curr_add_line1 + ', ' + tp_curr_add_line2 + ', ' + tp_curr_add_line3)," _
                        & " tp_curr_postcode, tp_curr_city, tp_curr_state," _
                        & " tp_tel1, tp_tel2, tp_mobile1, tp_mobile2," _
                        & " (tp_employer_no2 + tp_employer_no3)," _
                        & " tp_email, tp_bank, tp_bank_acc," _
                        & " tp_hw_name, tp_hw_ref_no_prefix, tp_hw_ref_no1, tp_hw_ref_no2, tp_hw_ref_no3," _
                        & " (tp_hw_ic_new1 + tp_hw_ic_new2 + tp_hw_ic_new3), tp_hw_ic_old," _
                        & " tp_hw_police_no, tp_hw_army_no, tp_hw_passport_no, tp_assessmenton" _
                        & " from taxp_profile where (tp_ref_no1 + tp_ref_no2 + tp_ref_no3)=?", prmOledb)

            dr = datHandler.GetDataReader("select tp_last_passport_no, TP_WORKER_APPROVEDATE, TP_COM_ADD_STATUS from taxp_profile2 where" _
                        & " tp_ref_no= '" & pdfForm.GetRefNo & "'")

            If ds.Tables(0).Rows.Count > 0 Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) Then
                    strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString(), 28)
                    If Not String.IsNullOrEmpty(strArray(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "I_1[0]", strArray(0).ToUpper)
                        pdfFormFields.SetField(pdfFieldPath & "I_2[0]", strArray(1).ToUpper)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(1)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(1).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "II_1[0]", ds.Tables(0).Rows(0).Item(1).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(2)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(2).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "II_2[0]", ds.Tables(0).Rows(0).Item(2).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(3)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(3).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "III[0]", ds.Tables(0).Rows(0).Item(3).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(4)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(4).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "IV[0]", ds.Tables(0).Rows(0).Item(4).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(5)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(5).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "V[0]", ds.Tables(0).Rows(0).Item(5).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(6)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(6).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "VI[0]", ds.Tables(0).Rows(0).Item(6).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(7)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(7).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "VII[0]", ds.Tables(0).Rows(0).Item(7).ToString)
                    End If
                End If
                If dr.Read() Then
                    'lyeyc
                    If Not IsDBNull(dr("TP_COM_ADD_STATUS")) Then
                        If Not String.IsNullOrEmpty(dr("TP_COM_ADD_STATUS")) Then
                            If dr("TP_COM_ADD_STATUS") = "1" Then
                                pdfFormFields.SetField(pdfFieldPath & "A7_7", "X")
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "A7_7", "")
                            End If
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "A7_7", "")
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A7_7", "")
                    End If
                    'lyeyc (end)
                    If Not IsDBNull(dr("tp_last_passport_no")) Then
                        If Not String.IsNullOrEmpty(dr("tp_last_passport_no").ToString) Then
                            pdfFormFields.SetField(pdfFieldPath & "VIII[0]", dr("tp_last_passport_no").ToString)
                        End If
                    End If
                    'weihong
                    If Not IsDBNull(dr("TP_WORKER_APPROVEDATE")) Then
                        pdfFormFields.SetField(pdfFieldPath & "A7[0]", "1")
                        pdfFormFields.SetField(pdfFieldPath & "A7a[0]", FormatDate(dr("TP_WORKER_APPROVEDATE")))
                        ' pdfFormFields.SetField(pdfFieldPath & "A4[0]", FormatDate(ds.Tables(0).Rows(0).Item(11)))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A7[0]", "2")
                    End If
                    'weihong

                End If
                dr.Close()

                'Initialise
                pdfFormFields.SetField(pdfFieldPath & "IX_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "IX_2[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "IX_3[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "IX_4[0]", "")

                Select Case (GetStatusOfTax())
                    Case "REPAYABLE"
                        pdfFormFields.SetField(pdfFieldPath & "IX_1[0]", "X")
                        pdfFormFields.SetField(pdfFieldPath & "IX_2[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "IX_3[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "IX_4[0]", "")
                    Case "EXCESS"
                        pdfFormFields.SetField(pdfFieldPath & "IX_1[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "IX_2[0]", "X")
                        pdfFormFields.SetField(pdfFieldPath & "IX_3[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "IX_4[0]", "")
                    Case "BALANCE"
                        pdfFormFields.SetField(pdfFieldPath & "IX_1[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "IX_2[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "IX_3[0]", "X")
                        pdfFormFields.SetField(pdfFieldPath & "IX_4[0]", "")
                    Case "NIL"
                        pdfFormFields.SetField(pdfFieldPath & "IX_1[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "IX_2[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "IX_3[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "IX_4[0]", "X")
                End Select

                ' ==== PART A ==== "
                ReDim strArray(2)

                If Not IsDBNull(ds.Tables(0).Rows(0).Item(8)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(8).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A1[0]", ds.Tables(0).Rows(0).Item(8).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(9)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(9).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A2[0]", ds.Tables(0).Rows(0).Item(9).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(10)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(10).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A3[0]", ds.Tables(0).Rows(0).Item(10).ToString)
                        If ds.Tables(0).Rows(0).Item(10).ToString = "2" Then
                            If Not IsDBNull(ds.Tables(0).Rows(0).Item(11)) Then
                                pdfFormFields.SetField(pdfFieldPath & "A4[0]", FormatDate(ds.Tables(0).Rows(0).Item(11)))
                            End If
                        ElseIf ds.Tables(0).Rows(0).Item(10).ToString = "3" Or ds.Tables(0).Rows(0).Item(10).ToString = "4" Then
                            If Not IsDBNull(ds.Tables(0).Rows(0).Item(12)) Then
                                pdfFormFields.SetField(pdfFieldPath & "A4[0]", FormatDate(ds.Tables(0).Rows(0).Item(12)))
                            End If
                        End If
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(13)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(13).ToString) Then
                        If ds.Tables(0).Rows(0).Item(13).ToString = "1" Then
                            If ds.Tables(0).Rows(0).Item(37).ToString = "1" Then
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "1")
                            ElseIf ds.Tables(0).Rows(0).Item(37).ToString = "2" Then
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "2")
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "")
                            End If
                        ElseIf ds.Tables(0).Rows(0).Item(13).ToString = "2" Then
                            pdfFormFields.SetField(pdfFieldPath & "A5[0]", "3")
                        ElseIf ds.Tables(0).Rows(0).Item(13).ToString = "3" Then
                            'weihong
                            If ds.Tables(0).Rows(0).Item(10).ToString = "2" Then
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "4")
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "5")
                            End If
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "A5[0]", "")
                        End If
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(14)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(14).ToString) Then
                        If ds.Tables(0).Rows(0).Item(14).ToString = "1" Then
                            pdfFormFields.SetField(pdfFieldPath & "A6[0]", "1")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "A6[0]", "2")
                        End If
                    End If
                End If

                If Not IsDBNull(ds.Tables(0).Rows(0).Item(15)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(15).ToString) Then
                        strArray = SplitText(ds.Tables(0).Rows(0).Item(15).ToString().Replace(",,", ",").Replace(", ,", ","), 26)
                        If Not String.IsNullOrEmpty(strArray(0)) Then
                            pdfFormFields.SetField(pdfFieldPath & "A7_1[0]", strArray(0).ToUpper)
                            pdfFormFields.SetField(pdfFieldPath & "A7_2[0]", strArray(1).ToUpper)
                            pdfFormFields.SetField(pdfFieldPath & "A7_3[0]", strArray(2).ToUpper)
                        End If
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(16)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(16).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A7_4[0]", ds.Tables(0).Rows(0).Item(16).ToString.ToUpper)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(17)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(17).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A7_5[0]", ds.Tables(0).Rows(0).Item(17).ToString.ToUpper)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(18)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(18).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A7_6[0]", ds.Tables(0).Rows(0).Item(18).ToString.ToUpper)
                    End If
                End If

                pdfFormFields.SetField(pdfFieldPath & "A8[0]", FormatPhoneNumber( _
                                            ds.Tables(0).Rows(0).Item(19).ToString, _
                                            ds.Tables(0).Rows(0).Item(20).ToString, _
                                            ds.Tables(0).Rows(0).Item(21).ToString, _
                                            ds.Tables(0).Rows(0).Item(22).ToString))

                If Not IsDBNull(ds.Tables(0).Rows(0).Item(23)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(23).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A9[0]", ds.Tables(0).Rows(0).Item(23).ToString)
                    End If
                End If

                ReDim strArray(0)
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(24)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(24).ToString) Then
                        strArray = SplitText(ds.Tables(0).Rows(0).Item(24).ToString, 28)
                        'pdfFormFields.SetField(pdfFieldPath & "A10[0]", ds.Tables(0).Rows(0).Item(24).ToString)
                        pdfFormFields.SetField(pdfFieldPath & "A10[0]", strArray(0))
                    End If
                End If

                ReDim strArray(0)
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(25)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(25).ToString.ToUpper) Then
                        strArray = SplitText(ds.Tables(0).Rows(0).Item(25).ToString, 28)
                        pdfFormFields.SetField(pdfFieldPath & "A11[0]", strArray(0))
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(26).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(26).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A12[0]", ds.Tables(0).Rows(0).Item(26).ToString)
                    End If
                End If

                ' ==== PART B ==== "
                ReDim strArray(1)
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(27)) Then
                    strArray = SplitText(ds.Tables(0).Rows(0).Item(27).ToString(), 28)
                    If Not String.IsNullOrEmpty(strArray(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "B1_1[0]", strArray(0).ToUpper)
                        pdfFormFields.SetField(pdfFieldPath & "B1_2[0]", strArray(1).ToUpper)
                    End If
                End If
                'lyeyc
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(28).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(28).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B2_1[0]", ds.Tables(0).Rows(0).Item(28).ToString)
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "B2_1[0]", "")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "B2_1[0]", "")
                End If
                'lyeyc (end)
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(29).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(29).ToString) Then
                        strHWIC = strHWIC & ds.Tables(0).Rows(0).Item(29).ToString
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(30).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(30).ToString) Then
                        strHWIC = strHWIC & ds.Tables(0).Rows(0).Item(30).ToString
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(31).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(31).ToString) Then
                        strHWIC = strHWIC & ds.Tables(0).Rows(0).Item(31).ToString
                    End If
                End If
                If Not String.IsNullOrEmpty(strHWIC) Then
                    pdfFormFields.SetField(pdfFieldPath & "B2_2[0]", strHWIC.ToString)
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(32).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(32).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B3[0]", ds.Tables(0).Rows(0).Item(32).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(33).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(33).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B4[0]", ds.Tables(0).Rows(0).Item(33).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(34).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(34).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B5[0]", ds.Tables(0).Rows(0).Item(34).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(35).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(35).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B6[0]", ds.Tables(0).Rows(0).Item(35).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(36).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(36).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B7[0]", ds.Tables(0).Rows(0).Item(36).ToString)
                    End If
                End If
            End If
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
        Dim dblTotalValue As Double = 0
        Dim dblTotalValue1 As Double = 0
        Dim intCounter As Integer = 1

        Try
            ' prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page4[0]."


            ' ==== PART C ==== "
            'C1 - C7
            dr = datHandler.GetDataReader("Select TC_KEY, TC_STATUTORY_INCOME, TC_BUSINESSLOSS_BF, TC_AGGREGATE_BUS_INCOME," _
                                & " TC_EMPLOYMENT_INCOME, TC_DIVIDEND, (cdbl(TC_INTEREST) + cdbl(TC_DISCOUNT)), " _
                                & " (cdbl(TC_RENTAL_ROYALTY)+cdbl(TC_PREMIUM)), TC_PENSION_AND_ETC," _
                                & " (cdbl(TC_OTHER_GAIN_PROFIT) + cdbl(TC_SEC4A)), TC_ADDITION_43," _
                                & " TC_AGGREGATE_OTHER_SRC, TC_AGGREGATE_INCOME, TC_BUSINESSLOSS_CY," _
                                & " TC_TOTAL1, TC_4, TC_3, TC_TOTAL_INCOME_2, TC_INCOME_TRANSFER_FROM_HW" _
                                & " from tax_computation where" _
                                & " tc_ref_no ='" & pdfForm.GetRefNo & "' and tc_ya ='" & pdfForm.GetYA & "'")

            If dr.Read() Then

                If Not IsDBNull(dr("TC_EMPLOYMENT_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C1[0]", FormatFixedAmount(dr("TC_EMPLOYMENT_INCOME")))
                End If
                If Not IsDBNull(dr("TC_DIVIDEND")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C2[0]", FormatFixedAmount(dr("TC_DIVIDEND")))
                End If
                If Not IsDBNull(dr.Item(6)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C3[0]", FormatFixedAmount(dr.Item(6).ToString))
                End If
                If Not IsDBNull(dr.Item(7)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C4[0]", FormatFixedAmount(dr.Item(7).ToString))
                End If
                If Not IsDBNull(dr("TC_PENSION_AND_ETC")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C5[0]", FormatFixedAmount(dr("TC_PENSION_AND_ETC")))
                End If
                If Not IsDBNull(dr.Item(9)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C6[0]", FormatFixedAmount(dr.Item(9).ToString))
                End If
                If Not IsDBNull(dr("TC_AGGREGATE_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C7[0]", FormatFixedAmount(dr("TC_AGGREGATE_INCOME")))
                End If



                'C8 - 'C15
                dr2 = datHandler.GetDataReader("select TCG_KEY, TCG_AMOUNT" _
                                        & " from tax_gifts where" _
                                        & " tc_key =" & dr("TC_KEY"))
                Do While dr2.Read()
                    If Not IsDBNull(dr2("TCG_KEY")) Then

                        Select Case dr2("TCG_KEY")
                            Case "9"
                                pdfFormFields.SetField(pdfFieldPath & "C8[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                            Case "1"
                                pdfFormFields.SetField(pdfFieldPath & "C8A[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "7"
                                pdfFormFields.SetField(pdfFieldPath & "C9[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "8"
                                pdfFormFields.SetField(pdfFieldPath & "C10_1[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "2"
                                pdfFormFields.SetField(pdfFieldPath & "C11[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                            Case "3"
                                pdfFormFields.SetField(pdfFieldPath & "C12[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                            Case "4"
                                pdfFormFields.SetField(pdfFieldPath & "C13[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                            Case "5"
                                pdfFormFields.SetField(pdfFieldPath & "C14[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                            Case "6"
                                pdfFormFields.SetField(pdfFieldPath & "C15[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))

                        End Select
                    End If
                Loop
                dr2.Close()

                'C10 Total restrict to 7% of C7
                If Not IsDBNull(dr("TC_AGGREGATE_INCOME")) Then
                    If Not String.IsNullOrEmpty(dr("TC_AGGREGATE_INCOME").ToString) Then
                        If dblTotalValue >= CDbl(dr("TC_AGGREGATE_INCOME")) * 0.07 Then
                            pdfFormFields.SetField(pdfFieldPath & "C10_2[0]", FormatFixedAmount((CDbl(dr("TC_AGGREGATE_INCOME")) * 0.07).ToString))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "C10_2[0]", FormatFixedAmount(dblTotalValue.ToString))
                        End If
                    End If
                End If

                'C16 - C17_2
                If Not IsDBNull(dr("TC_3")) And Not IsDBNull(dr("TC_4")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C16[0]", FormatFloatingAmount(CDbl(dr("TC_3")) + CDbl(dr("TC_4")), False))
                ElseIf Not IsDBNull(dr("TC_4")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C16[0]", dr("TC_4"))
                ElseIf Not IsDBNull(dr("TC_3")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C16[0]", dr("TC_3"))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "C16[0]", "0")
                End If

                If Not IsDBNull(dr("TC_INCOME_TRANSFER_FROM_HW")) Then
                    If Not String.IsNullOrEmpty(dr("TC_INCOME_TRANSFER_FROM_HW").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "C17_2[0]", FormatFixedAmount(dr("TC_INCOME_TRANSFER_FROM_HW")))
                    End If
                End If

                'C17_1
                'NGOHCS B2010.2
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
                'NGOHCS B2010.2 END

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
                                        pdfFormFields.SetField(pdfFieldPath & "C17_1[0]", "1")
                                    Else
                                        pdfFormFields.SetField(pdfFieldPath & "C17_1[0]", "2")
                                    End If
                                Else
                                    pdfFormFields.SetField(pdfFieldPath & "C17_1[0]", "")
                                End If
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "C17_1[0]", "")
                            End If
                        End If
                    End If

                    'C18
                    'NGOHCS B2010.2
                    'If Not IsDBNull(dr("TC_3")) And Not IsDBNull(dr("TC_4")) Then
                    '    pdfFormFields.SetField(pdfFieldPath & "C18[0]", (CDbl(dr("TC_3")) + CDbl(dr("TC_4")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))))
                    'ElseIf Not IsDBNull(dr("TC_4")) Then
                    '    pdfFormFields.SetField(pdfFieldPath & "C18[0]", (CDbl(dr("TC_4")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))))
                    'ElseIf Not IsDBNull(dr("TC_3")) Then
                    '    pdfFormFields.SetField(pdfFieldPath & "C18[0]", (CDbl(dr("TC_3")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))))
                    'Else
                    '    pdfFormFields.SetField(pdfFieldPath & "C18[0]", CDbl(dr("TC_INCOME_TRANSFER_FROM_HW")))
                    'End If

                    If Not IsDBNull(dr2("TP_STATUS")) Then
                        If dr2("TP_STATUS") = "1" Then
                            pdfFormFields.SetField(pdfFieldPath & "C18[0]", "0")
                        ElseIf dr2("TP_STATUS") = "2" And CDbl(pdfForm.GetYA) >= 2007 And dr2("TP_TYPE_ASSESSMENT") = "3" Then
                            pdfFormFields.SetField(pdfFieldPath & "C18[0]", "0")
                        ElseIf dr2("TP_STATUS") = "2" And CDbl(pdfForm.GetYA) = 2006 And dr2("TP_TYPE_ASSESSMENT") = "1" And CDbl(dr("TC_INCOME_TRANSFER_FROM_HW")) = 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "C18[0]", "0")
                        ElseIf (dr2("TP_GENDER") = "1" And dr2("TP_STATUS") = "2" And dr2("TP_TYPE_ASSESSMENT") = "1" And _
                            dr2("TP_ASSESSMENTON") = "1") Or (dr2("TP_GENDER") = "2" And dr2("TP_STATUS") = "2" And dr2("TP_TYPE_ASSESSMENT") = "1" And dr2("TP_ASSESSMENTON") = "2") Then
                            pdfFormFields.SetField(pdfFieldPath & "C18[0]", FormatFixedAmount((CDbl(dr("TC_TOTAL_INCOME_2")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))).ToString))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "C18[0]", "0")
                        End If
                    End If
                End If
                dr2.Close()

                ' ==== PART D ==== "
                ' NOTE: Element D11, D14 was filled up in part E

                dblTotalValue = 0
                dr2 = datHandler.GetDataReader("select tcc_key, tcc_amount from tax_relief where " _
                                   & " tc_key =" & dr("TC_KEY") & " order by tcc_key")

                'D1 - D14
                While dr2.Read()
                    If Not IsDBNull(dr2("tcc_key")) And Not IsDBNull(dr2("tcc_amount")) Then
                        If Not String.IsNullOrEmpty(dr2("tcc_amount")) Then

                            Select Case dr2("tcc_key")
                                Case 2
                                    pdfFormFields.SetField(pdfFieldPath & "D2[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 3
                                    pdfFormFields.SetField(pdfFieldPath & "D3[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 4
                                    pdfFormFields.SetField(pdfFieldPath & "D4[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 5
                                    pdfFormFields.SetField(pdfFieldPath & "D5[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 6
                                    pdfFormFields.SetField(pdfFieldPath & "D6[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                    dblTotalValue = dblTotalValue + CDbl(dr2("tcc_amount"))
                                Case 7
                                    pdfFormFields.SetField(pdfFieldPath & "D7_1[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                    dblTotalValue = dblTotalValue + CDbl(dr2("tcc_amount"))
                                Case 8
                                    pdfFormFields.SetField(pdfFieldPath & "D8[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 9
                                    pdfFormFields.SetField(pdfFieldPath & "D8D[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 10
                                    pdfFormFields.SetField(pdfFieldPath & "D9[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 13
                                    pdfFormFields.SetField(pdfFieldPath & "D10[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 14
                                    pdfFormFields.SetField(pdfFieldPath & "D11a_5[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 15
                                    pdfFormFields.SetField(pdfFieldPath & "D11b_5[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 16
                                    pdfFormFields.SetField(pdfFieldPath & "D11c_5[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 17
                                    pdfFormFields.SetField(pdfFieldPath & "D12[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                    'weihong
                                    dblTotalValue1 = dblTotalValue1 + CDbl(dr2("tcc_amount"))
                                Case 18
                                    pdfFormFields.SetField(pdfFieldPath & "D13[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 21
                                    pdfFormFields.SetField(pdfFieldPath & "D8A[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 22
                                    pdfFormFields.SetField(pdfFieldPath & "D8B[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 23
                                    pdfFormFields.SetField(pdfFieldPath & "D8C[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                    'weihong
                                Case 24
                                    'Yuran langganan perkhidmatan internet jalur lebar
                                    pdfFormFields.SetField(pdfFieldPath & "D12_I[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 25
                                    pdfFormFields.SetField(pdfFieldPath & "D18[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                    dblTotalValue1 = dblTotalValue1 + CDbl(dr2("tcc_amount"))
                            End Select
                        End If
                    End If
                End While
                pdfFormFields.SetField(pdfFieldPath & "D7_2[0]", FormatFixedAmount(dblTotalValue.ToString))
                pdfFormFields.SetField(pdfFieldPath & "D17[0]", FormatFixedAmount(dblTotalValue1.ToString))
                'pdfFormFields.SetField(pdfFieldPath & "D17[0]", FormatFixedAmount(dblTotalValue.ToString))
                dr2.Close()
            End If
            dr.Close()

            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama4", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj4", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub Page3()

        Dim pdfFieldPath As String
        Dim dr As OleDbDataReader = Nothing
        Dim dr2 As OleDbDataReader = Nothing
        Dim dblTotalValue As Double = 0
        Dim boolHasHusbandWife As Boolean = False
        Dim strHWRefNo As String = ""
        Dim intCounter(2) As Integer
        Dim intArrayChild50(5) As Integer
        Dim intArrayChild100(5) As Integer

        Try
            ' prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page5[0]."

            ' ==== Part E ==== '
            'E1 - E15 exclude E4 - E7
            dr = datHandler.GetDataReader("Select TC_KEY, TC_RELIEF, TC_CHARGEABLE_INCOME, TC_TAX_FIRST_INCOME, TC_TAX_FIRST_TAX," _
                                & " TC_TAX_BALANCE_INCOME, TC_TAX_BALANCE_RATE, TC_TAX_BALANCE_TAX, TC_TOTAL_INCOME_TAX," _
                                & " TC_REBATES, TC_INCOME_TAX_CHARGED, TC_SEC110_DIVIDEND,TC_SEC110_OTHERS, TC_1," _
                                & " TC_2, TC_TAX_PAYABLE, TC_TAX_REPAYMENT, TC_TAX_SCH1_INCOME, TC_TAX_SCH1_TAX" _
                                & " From TAX_COMPUTATION Where" _
                                & " TC_REF_NO= '" & pdfForm.GetRefNo & "' and TC_YA= '" & pdfForm.GetYA & "'")

            If dr.Read() Then
                'weihong
                If Not IsDBNull(dr("TC_TAX_SCH1_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E2_1[0]", FormatFixedAmount(dr("TC_TAX_SCH1_INCOME").ToString))
                End If
                If Not IsDBNull(dr("TC_TAX_SCH1_TAX")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E2_2[0]", FormatFloatingAmount(dr("TC_TAX_SCH1_TAX").ToString, True))
                End If

                'D14
                If Not IsDBNull(dr("TC_RELIEF")) Then
                    pdfFormFields.SetField(pdfFieldPath & "D14[0]", FormatFixedAmount(dr("TC_RELIEF").ToString))
                End If
                If Not IsDBNull(dr("TC_CHARGEABLE_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E1[0]", FormatFixedAmount(dr("TC_CHARGEABLE_INCOME").ToString))
                End If
                If Not IsDBNull(dr("TC_TAX_FIRST_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E2a_1[0]", FormatFixedAmount(dr("TC_TAX_FIRST_INCOME").ToString))
                End If
                If Not IsDBNull(dr("TC_TAX_FIRST_TAX")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E2a_2[0]", FormatFloatingAmount(dr("TC_TAX_FIRST_TAX").ToString, True))
                End If
                If Not IsDBNull(dr("TC_TAX_BALANCE_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E2b_1[0]", FormatFixedAmount(dr("TC_TAX_BALANCE_INCOME").ToString))
                End If
                If Not IsDBNull(dr("TC_TAX_BALANCE_RATE")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E2b_2[0]", FormatFixedAmount(dr("TC_TAX_BALANCE_RATE").ToString))
                End If
                If Not IsDBNull(dr("TC_TAX_BALANCE_TAX")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E2b_3[0]", FormatFloatingAmount(dr("TC_TAX_BALANCE_TAX").ToString, True))
                End If
                If Not IsDBNull(dr("TC_TOTAL_INCOME_TAX")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E3[0]", FormatFloatingAmount(dr("TC_TOTAL_INCOME_TAX").ToString, True))
                End If
                If Not IsDBNull(dr("TC_REBATES")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E8[0]", FormatFloatingAmount(dr("TC_REBATES").ToString, True))
                End If
                If Not IsDBNull(dr("TC_INCOME_TAX_CHARGED")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E9[0]", FormatFloatingAmount(dr("TC_INCOME_TAX_CHARGED").ToString, True))
                End If
                If Not IsDBNull(dr("TC_SEC110_DIVIDEND")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E10[0]", FormatFloatingAmount(dr("TC_SEC110_DIVIDEND").ToString, True))
                End If
                If Not IsDBNull(dr("TC_SEC110_OTHERS")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E11[0]", FormatFloatingAmount(dr("TC_SEC110_OTHERS").ToString, True))
                End If
                If Not IsDBNull(dr("TC_1")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E12[0]", FormatFloatingAmount(dr("TC_1").ToString, True))
                End If
                If Not IsDBNull(dr("TC_2")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E13[0]", FormatFloatingAmount(dr("TC_2").ToString, True))
                End If
                If Not IsDBNull(dr("TC_TAX_PAYABLE")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E14[0]", FormatFloatingAmount(dr("TC_TAX_PAYABLE").ToString, True))
                End If
                If Not IsDBNull(dr("TC_TAX_REPAYMENT")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E15[0]", FormatFloatingAmount(dr("TC_TAX_REPAYMENT").ToString, True))
                End If

                'E4 - E7
                dr2 = datHandler.GetDataReader("SELECT TCR_KEY, TCR_AMOUNT FROM [TAX_REBATE] WHERE [TC_KEY]= " & dr("TC_KEY"))
                While dr2.Read()
                    If Not IsDBNull(dr2("TCR_KEY")) Then
                        If Not String.IsNullOrEmpty(dr2("TCR_KEY").ToString) Then
                            Select Case dr2("TCR_KEY")
                                Case 1
                                    pdfFormFields.SetField(pdfFieldPath & "E4[0]", FormatFixedAmount(dr2("TCR_AMOUNT").ToString))
                                Case 2
                                    pdfFormFields.SetField(pdfFieldPath & "E5[0]", FormatFixedAmount(dr2("TCR_AMOUNT").ToString))
                                Case 3
                                    pdfFormFields.SetField(pdfFieldPath & "E6[0]", FormatFloatingAmount(FormatNumber(CDbl(dr2("TCR_AMOUNT")), 2).ToString, True))
                                Case 5
                                    pdfFormFields.SetField(pdfFieldPath & "E7[0]", FormatFloatingAmount(FormatNumber(CDbl(dr2("TCR_AMOUNT")), 2).ToString, True))
                            End Select
                        End If
                    End If
                End While
                dr2.Close()


                'D11 Relief Child

                dblTotalValue = 0
                strHWRefNo = ""
                boolHasHusbandWife = False


                ReDim intCounter(2)
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
                                    ElseIf CDbl(Trim(dr2("TCC_100"))) = 4000 Then
                                        intArrayChild100(3) = intArrayChild100(3) + 1
                                    End If
                                End If
                                If Not IsDBNull(dr2("TCC_50")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_50"))) And Not Trim(dr2("TCC_50")) = "" Then
                                    If CDbl(Trim(dr2("TCC_50"))) = 500 Then
                                        intArrayChild50(2) = intArrayChild50(2) + 1
                                    ElseIf CDbl(Trim(dr2("TCC_50"))) = 2000 Then
                                        intArrayChild50(3) = intArrayChild50(3) + 1
                                    End If
                                End If
                            Case 16
                                If Not IsDBNull(dr2("TCC_100")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_100"))) And Not Trim(dr2("TCC_100")) = "" Then
                                    If CDbl(Trim(dr2("TCC_100"))) = 5000 Then
                                        intArrayChild100(4) = intArrayChild100(4) + 1
                                    ElseIf CDbl(Trim(dr2("TCC_100"))) = 9000 Then
                                        intArrayChild100(5) = intArrayChild100(5) + 1
                                    End If
                                End If
                                If Not IsDBNull(dr2("TCC_50")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_50"))) And Not Trim(dr2("TCC_50")) = "" Then
                                    If CDbl(Trim(dr2("TCC_50"))) = 2500 Then
                                        intArrayChild50(4) = intArrayChild50(4) + 1
                                    ElseIf CDbl(Trim(dr2("TCC_50"))) = 4500 Then
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
                            pdfFormFields.SetField(pdfFieldPath & "D11a_1[0]", intArrayChild100(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "D11a_2[0]", intArrayChild100(i) * 1000)
                            pdfFormFields.SetField(pdfFieldPath & "D11a_3[0]", intArrayChild50(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "D11a_4[0]", intArrayChild50(i) * 500)
                        Case 2
                            pdfFormFields.SetField(pdfFieldPath & "D11b1_1[0]", intArrayChild100(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "D11b1_2[0]", intArrayChild100(i) * 1000)
                            pdfFormFields.SetField(pdfFieldPath & "D11b1_3[0]", intArrayChild50(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "D11b1_4[0]", intArrayChild50(i) * 500)
                        Case 3
                            pdfFormFields.SetField(pdfFieldPath & "D11b2_1[0]", intArrayChild100(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "D11b2_2[0]", intArrayChild100(i) * 4000)
                            pdfFormFields.SetField(pdfFieldPath & "D11b2_3[0]", intArrayChild50(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "D11b2_4[0]", intArrayChild50(i) * 2000)
                        Case 4
                            pdfFormFields.SetField(pdfFieldPath & "D11c1_1[0]", intArrayChild100(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "D11c1_2[0]", intArrayChild100(i) * 5000)
                            pdfFormFields.SetField(pdfFieldPath & "D11c1_3[0]", intArrayChild50(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "D11c1_4[0]", intArrayChild50(i) * 2500)
                        Case 5
                            pdfFormFields.SetField(pdfFieldPath & "D11c2_1[0]", intArrayChild100(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "D11c2_2[0]", intArrayChild100(i) * 9000)
                            pdfFormFields.SetField(pdfFieldPath & "D11c2_3[0]", intArrayChild50(i).ToString)
                            pdfFormFields.SetField(pdfFieldPath & "D11c2_4[0]", intArrayChild50(i) * 4500)
                    End Select
                    intCounter(0) = intCounter(0) + intArrayChild50(i) + intArrayChild100(i)
                Next
            End If
            dr.Close()
            'D11 Bilangan anak dituntut oleh diri sendiri
            pdfFormFields.SetField(pdfFieldPath & "D11_2[0]", intCounter(0).ToString)


            dr = datHandler.GetDataReader("Select tp_hw_ref_no1, tp_hw_ref_no2, tp_hw_ref_no3" _
                            & " from taxp_profile where (tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then strHWRefNo = strHWRefNo + dr.Item(0).ToString
                If Not IsDBNull(dr.Item(1)) Then strHWRefNo = strHWRefNo + dr.Item(1).ToString
                If Not IsDBNull(dr.Item(2)) Then strHWRefNo = strHWRefNo + dr.Item(2).ToString
                boolHasHusbandWife = True
            End If
            dr.Close()

            If boolHasHusbandWife Then

                For i As Integer = 1 To 5
                    intArrayChild50(i) = 0
                    intArrayChild100(i) = 0
                Next

                dr = datHandler.GetDataReader("Select TC_KEY From TAX_COMPUTATION Where" _
                            & " TC_REF_NO='" & strHWRefNo & "' And TC_YA='" & pdfForm.GetYA & "'")

                If dr.Read() Then
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
                                        ElseIf CDbl(Trim(dr2("TCC_100"))) = 4000 Then
                                            intArrayChild100(3) = intArrayChild100(3) + 1
                                        End If
                                    End If
                                    If Not IsDBNull(dr2("TCC_50")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_50"))) And Not Trim(dr2("TCC_50")) = "" Then
                                        If CDbl(Trim(dr2("TCC_50"))) = 500 Then
                                            intArrayChild50(2) = intArrayChild50(2) + 1
                                        ElseIf CDbl(Trim(dr2("TCC_50"))) = 2000 Then
                                            intArrayChild50(3) = intArrayChild50(3) + 1
                                        End If
                                    End If
                                Case 16
                                    If Not IsDBNull(dr2("TCC_100")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_100"))) And Not Trim(dr2("TCC_100")) = "" Then
                                        If CDbl(Trim(dr2("TCC_100"))) = 5000 Then
                                            intArrayChild100(4) = intArrayChild100(4) + 1
                                        ElseIf CDbl(Trim(dr2("TCC_100"))) = 9000 Then
                                            intArrayChild100(5) = intArrayChild100(5) + 1
                                        End If
                                    End If
                                    If Not IsDBNull(dr2("TCC_50")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_50"))) And Not Trim(dr2("TCC_50")) = "" Then
                                        If CDbl(Trim(dr2("TCC_50"))) = 2500 Then
                                            intArrayChild50(4) = intArrayChild50(4) + 1
                                        ElseIf CDbl(Trim(dr2("TCC_50"))) = 4500 Then
                                            intArrayChild50(5) = intArrayChild50(5) + 1
                                        End If
                                    End If
                            End Select
                        End If
                    End While

                    'For i As Integer = 1 To 5
                    '    intCounter(1) = intCounter(1) + intArrayChild50(i) + intArrayChild100(i)
                    'Next

                    dr2.Close()
                End If
                dr.Close()

                dr = datHandler.GetDataReader("Select tp_hw_ref_no1, tp_hw_ref_no2, tp_hw_ref_no3" _
                        & " from taxp_profile_hw_others where tp_ref_no= '" & pdfForm.GetRefNo & "'")
                While dr.Read()
                    strHWRefNo = ""
                    If Not IsDBNull(dr.Item(0)) Then strHWRefNo = strHWRefNo + dr.Item(0).ToString
                    If Not IsDBNull(dr.Item(1)) Then strHWRefNo = strHWRefNo + dr.Item(1).ToString
                    If Not IsDBNull(dr.Item(2)) Then strHWRefNo = strHWRefNo + dr.Item(2).ToString
                    'boolHasHusbandWife = True

                    dr2 = datHandler.GetDataReader("SELECT TCC_KEY, TCC_100, TCC_50 FROM [TAX_RELIEF_CHILD] WHERE [TC_KEY] in " & _
                            "(Select TC_KEY From TAX_COMPUTATION Where" & _
                            " TC_REF_NO='" & strHWRefNo & "' And TC_YA='" & pdfForm.GetYA & "') order by  [TCC_KEY]")

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
                                        ElseIf CDbl(Trim(dr2("TCC_100"))) = 4000 Then
                                            intArrayChild100(3) = intArrayChild100(3) + 1
                                        End If
                                    End If
                                    If Not IsDBNull(dr2("TCC_50")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_50"))) And Not Trim(dr2("TCC_50")) = "" Then
                                        If CDbl(Trim(dr2("TCC_50"))) = 500 Then
                                            intArrayChild50(2) = intArrayChild50(2) + 1
                                        ElseIf CDbl(Trim(dr2("TCC_50"))) = 2000 Then
                                            intArrayChild50(3) = intArrayChild50(3) + 1
                                        End If
                                    End If
                                Case 16
                                    If Not IsDBNull(dr2("TCC_100")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_100"))) And Not Trim(dr2("TCC_100")) = "" Then
                                        If CDbl(Trim(dr2("TCC_100"))) = 5000 Then
                                            intArrayChild100(4) = intArrayChild100(4) + 1
                                        ElseIf CDbl(Trim(dr2("TCC_100"))) = 9000 Then
                                            intArrayChild100(5) = intArrayChild100(5) + 1
                                        End If
                                    End If
                                    If Not IsDBNull(dr2("TCC_50")) And Not String.IsNullOrEmpty(Trim(dr2("TCC_50"))) And Not Trim(dr2("TCC_50")) = "" Then
                                        If CDbl(Trim(dr2("TCC_50"))) = 2500 Then
                                            intArrayChild50(4) = intArrayChild50(4) + 1
                                        ElseIf CDbl(Trim(dr2("TCC_50"))) = 4500 Then
                                            intArrayChild50(5) = intArrayChild50(5) + 1
                                        End If
                                    End If
                            End Select
                        End If
                    End While

                    dr2.Close()
                End While
                dr.Close()

                For i As Integer = 1 To 5
                    intCounter(1) = intCounter(1) + intArrayChild50(i) + intArrayChild100(i)
                Next
            End If

            'D11 Bilangan anak dituntut oleh suami / isteri
            pdfFormFields.SetField(pdfFieldPath & "D11_3[0]", intCounter(1).ToString)

            'D11 Bilangan anak layak mendapat pelepasan
            pdfFormFields.SetField(pdfFieldPath & "D11_1[0]", (intCounter(0) + intCounter(1)).ToString)


            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama5", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj5", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub Page4()

        Dim pdfFieldPath As String
        Dim strTempString As String = ""
        Dim dr As OleDbDataReader = Nothing
        Dim dr2 As OleDbDataReader = Nothing
        Dim dblTemp As Double = 0
        Dim dblTotalValue As Double = 0
        Dim dblTotalValue2 As Double = 0
        Dim intCounter As Integer = 1
        Dim strArray(1) As String


        Try
            ' prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page6[0]."


            ' === Part F === '
            dr = datHandler.GetDataReader("Select TC_TAX_PAYABLE, (cdbl(TC_INSTALLMENT_PAYMENT_SELF) + cdbl(TC_INSTALLMENT_PAYMENT_HW))," _
                                     & " TC_BALANCE_TAX_PAYABLE, TC_BALANCE_TAX_OVERPAID" _
                                     & " From TAX_COMPUTATION Where" _
                                     & " TC_REF_NO= '" & pdfForm.GetRefNo & "' and TC_YA= '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                If Not IsDBNull(dr("TC_TAX_PAYABLE")) Then
                    If Not String.IsNullOrEmpty(dr("TC_TAX_PAYABLE").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "F1[0]", FormatFloatingAmount(FormatNumber(CDbl(dr("TC_TAX_PAYABLE")), 2).ToString, True))
                    End If
                End If
                If Not IsDBNull(dr.Item(1)) Then
                    If Not String.IsNullOrEmpty(dr.Item(1).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "F2[0]", FormatFloatingAmount(FormatNumber(CDbl(dr.Item(1)), 2).ToString, True))
                    End If
                End If
                If Not IsDBNull(dr("TC_BALANCE_TAX_PAYABLE")) Then
                    If Not String.IsNullOrEmpty(dr("TC_BALANCE_TAX_PAYABLE").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "F3[0]", FormatFloatingAmount(FormatNumber(CDbl(dr("TC_BALANCE_TAX_PAYABLE")), 2).ToString, True))
                    End If
                End If
                If Not IsDBNull(dr("TC_BALANCE_TAX_OVERPAID")) Then
                    If Not String.IsNullOrEmpty(dr("TC_BALANCE_TAX_OVERPAID").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "F4[0]", FormatFloatingAmount(FormatNumber(CDbl(dr("TC_BALANCE_TAX_OVERPAID")), 2).ToString, True))
                    End If
                End If
            End If
            dr.Close()


            ' === Part G === '
            'G1 - G5
            intCounter = 1
            pdfFormFields.SetField(pdfFieldPath & "G1_1[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "G2_1[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "G3_1[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "G4_1[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "G5_1[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "G1_2[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "G2_2[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "G3_2[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "G4_2[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "G5_2[0]", "---")
            pdfFormFields.SetField(pdfFieldPath & "G1_3[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "G2_3[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "G3_3[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "G4_3[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "G5_3[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "G1_4[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "G2_4[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "G3_4[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "G4_4[0]", "0")
            pdfFormFields.SetField(pdfFieldPath & "G5_4[0]", "0")

            dr = datHandler.GetDataReader("Select * from preceding_year where py_ref_no= '" & pdfForm.GetRefNo & "' and py_ya= '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                dr2 = datHandler.GetDataReader("Select TOP 5 PY_INCOME_TYPE, PY_PAYMENT_YEAR, PY_AMOUNT, PY_EPF" _
                                                    & " From PRECEDING_YEAR_DETAIL Where" _
                                                    & " PY_KEY= " & dr("PY_KEY") & " Order By PY_DKEY")
                While dr2.Read()

                    If Not IsDBNull(dr2("PY_INCOME_TYPE")) Then
                        pdfFormFields.SetField(pdfFieldPath & "G" & intCounter.ToString & "_1[0]", dr2("PY_INCOME_TYPE").ToString.ToUpper)
                    End If
                    If Not IsDBNull(dr2("PY_PAYMENT_YEAR")) Then
                        pdfFormFields.SetField(pdfFieldPath & "G" & intCounter.ToString & "_2[0]", dr2("PY_PAYMENT_YEAR").ToString)
                    End If
                    If Not IsDBNull(dr2("PY_AMOUNT")) Then
                        pdfFormFields.SetField(pdfFieldPath & "G" & intCounter.ToString & "_3[0]", FormatFixedAmount(dr2("PY_AMOUNT").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "G" & intCounter.ToString & "_3[0]", "0")
                    End If
                    If Not IsDBNull(dr2("PY_EPF")) Then
                        pdfFormFields.SetField(pdfFieldPath & "G" & intCounter.ToString & "_4[0]", FormatFixedAmount(dr2("PY_EPF").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "G" & intCounter.ToString & "_4[0]", "0")

                    End If
                    intCounter = intCounter + 1
                End While
                dr2.Close()
            End If
            dr.Close()


            ' === Part H === '
            'H1 - H6
            dr = datHandler.GetDataReader("Select TP_ADM_NAME, (TP_ADM_IC_NEW1 + TP_ADM_IC_NEW2 + TP_ADM_IC_NEW3), TP_ADM_IC_OLD," _
                                     & " TP_ADM_POLICE_NO, TP_ADM_ARMY_NO, TP_ADM_PASSPORT_NO" _
                                     & " From TAXP_PROFILE Where" _
                                     & " (TP_REF_NO1 + TP_REF_NO2 + TP_REF_NO3)= '" & pdfForm.GetRefNo & "'")

            If dr.Read() Then

                If Not IsDBNull(dr("TP_ADM_NAME")) Then
                    If Not String.IsNullOrEmpty(dr("TP_ADM_NAME")) Then
                        strArray = SplitText(dr("TP_ADM_NAME").ToString, 28)
                        If Not String.IsNullOrEmpty(strArray(0)) Then
                            pdfFormFields.SetField(pdfFieldPath & "H1_1[0]", strArray(0).ToString.ToUpper)
                            pdfFormFields.SetField(pdfFieldPath & "H1_2[0]", strArray(1).ToString.ToUpper)
                        End If
                    End If
                End If

                If Not String.IsNullOrEmpty(dr.Item(1)) Then
                    pdfFormFields.SetField(pdfFieldPath & "H2[0]", dr.Item(1).ToString)
                End If
                If Not String.IsNullOrEmpty(dr("TP_ADM_IC_OLD")) Then 'weihong
                    pdfFormFields.SetField(pdfFieldPath & "H3[0]", dr("TP_ADM_IC_OLD").ToString)
                End If
                If Not String.IsNullOrEmpty(dr("TP_ADM_POLICE_NO")) Then 'weihong
                    pdfFormFields.SetField(pdfFieldPath & "H4[0]", dr("TP_ADM_POLICE_NO").ToString)
                End If
                If Not String.IsNullOrEmpty(dr("TP_ADM_POLICE_NO")) Then 'weihong
                    pdfFormFields.SetField(pdfFieldPath & "H5[0]", dr("TP_ADM_ARMY_NO").ToString)
                End If
                If Not String.IsNullOrEmpty(dr("TP_ADM_PASSPORT_NO")) Then 'weihong
                    pdfFormFields.SetField(pdfFieldPath & "H6[0]", dr("TP_ADM_PASSPORT_NO").ToString)
                End If
            End If
            dr.Close()


            ' === Part Akuan === '

            ReDim strArray(1)
            If Not String.IsNullOrEmpty(pdfForm.GetDeclarationReturn) Then
                pdfFormFields.SetField(pdfFieldPath & "Akuan3", pdfForm.GetDeclarationReturn)

                If pdfForm.GetDeclarationReturn = "1" Then
                    dr = datHandler.GetDataReader("SELECT * FROM [TAXP_PROFILE] WHERE (TP_REF_NO1 + TP_REF_NO2 + TP_REF_NO3) = '" & pdfForm.GetRefNo & "'")
                    If dr.Read() Then
                        If Not IsDBNull(dr("TP_NAME")) Then
                            If Not String.IsNullOrEmpty(dr("TP_NAME").ToString) Then
                                strArray = SplitText(dr("TP_NAME").ToString, 28)
                                If Not String.IsNullOrEmpty(strArray(0)) Then
                                    pdfFormFields.SetField(pdfFieldPath & "Akuan1_1", strArray(0).ToString.ToUpper)
                                    pdfFormFields.SetField(pdfFieldPath & "Akuan1_2", strArray(1).ToString.ToUpper)
                                End If
                            End If
                        End If
                        If Len(Trim(dr("TP_IC_NEW_1") + Trim(dr("TP_IC_NEW_2")) + Trim(dr("TP_IC_NEW_3")))) > 0 Then
                            strTempString = Trim(dr("TP_IC_NEW_1")) + Trim(dr("TP_IC_NEW_2")) + Trim(dr("TP_IC_NEW_3"))
                        ElseIf Len(Trim(dr("TP_PASSPORT_NO"))) > 0 Then
                            strTempString = (dr("TP_PASSPORT_NO"))
                        ElseIf Len(Trim(dr("TP_POLICE_NO"))) > 0 Then
                            strTempString = (dr("TP_POLICE_NO"))
                        ElseIf Len(Trim(dr("TP_ARMY_NO"))) > 0 Then
                            strTempString = (dr("TP_ARMY_NO"))
                        End If
                        pdfFormFields.SetField(pdfFieldPath & "Akuan2", strTempString)
                    End If
                Else
                    If Not String.IsNullOrEmpty(pdfForm.GetDeclarationBy) Then
                        strArray = SplitText(pdfForm.GetDeclarationBy, 28)
                        If Not String.IsNullOrEmpty(strArray(0)) Then
                            pdfFormFields.SetField(pdfFieldPath & "Akuan1_1", strArray(0).ToString.ToUpper)
                            pdfFormFields.SetField(pdfFieldPath & "Akuan1_2", strArray(1).ToString.ToUpper)
                        End If
                        If Not String.IsNullOrEmpty(pdfForm.GetDeclarationID.ToString) Then
                            pdfFormFields.SetField(pdfFieldPath & "Akuan2", pdfForm.GetDeclarationID)
                        End If
                    End If
                End If
            End If
            If Not String.IsNullOrEmpty(pdfForm.GetDeclarationDate) Then
                pdfFormFields.SetField(pdfFieldPath & "Akuan4", pdfForm.GetDeclarationDate)
            End If



            ' === Part Nyata === '

            ReDim strArray(0)
            dr = datHandler.GetDataReader("SELECT * FROM [TAXA_PROFILE] Where [TA_KEY] =" & pdfForm.GetTaxAgent)
            If dr.Read() Then
                If Not IsDBNull(dr("TA_CO_NAME")) Then
                    If Not String.IsNullOrEmpty(dr("TA_CO_NAME").ToString) Then
                        strArray = SplitText(dr("TA_CO_NAME").ToString, 29)
                        pdfFormFields.SetField(pdfFieldPath & "NyataA", strArray(0).ToUpper)
                    End If
                End If
                If Not IsDBNull(dr("TA_TEL_NO")) And Not IsDBNull(dr("TA_MOBILE")) Then
                    pdfFormFields.SetField(pdfFieldPath & "Nyatab", FormatPhoneNumber("", dr("TA_TEL_NO").ToString, "", dr("TA_MOBILE").ToString))
                End If
                If Not IsDBNull(dr("TA_LICENSE")) Then
                    pdfFormFields.SetField(pdfFieldPath & "Nyatac", dr("TA_LICENSE").ToString)
                End If
                pdfFormFields.SetField(pdfFieldPath & "NyataTarikh", FormatDate(Now))
            End If
            dr.Close()


            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama6", dr("tp_name").ToString.ToUpper)
                        pdfFormFields.SetField(pdfFieldPath & "Nama7", dr("tp_name").ToString.ToUpper)
                        pdfFormFields.SetField(pdfFieldPath & "Nama8", dr("tp_name").ToString.ToUpper)

                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj6", pdfForm.GetRefNo)
                    pdfFormFields.SetField(pdfFieldPath & "Ruj7", pdfForm.GetRefNo)
                    pdfFormFields.SetField(pdfFieldPath & "Ruj8", pdfForm.GetRefNo)


                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try


    End Sub

    Private Sub Page5()

        Dim pdfFieldPath As String = ""
        Dim strTempString As String = ""
        Dim dr As OleDbDataReader = Nothing


        Try
            ' === Part Slip === '
            'prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page7[0]."
            dr = datHandler.GetDataReader("Select * from taxp_profile where (tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read() Then
                If Not IsDBNull(dr("tp_ref_no_prefix")) And Not IsDBNull(dr("tp_ref_no1")) And Not IsDBNull(dr("tp_ref_no2")) And Not IsDBNull(dr("tp_ref_no3")) Then
                    If Not String.IsNullOrEmpty(dr("tp_ref_no_prefix").ToString & dr("tp_ref_no1").ToString & dr("tp_ref_no2").ToString & dr("tp_ref_no3").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Slip1", dr("tp_ref_no_prefix") & dr("tp_ref_no1") & dr("tp_ref_no2") & dr("tp_ref_no3"))
                    End If
                End If
                If Not IsDBNull(dr("TP_NAME")) Then
                    If Not String.IsNullOrEmpty(dr("TP_NAME").ToString) Then
                        strTempString = strTempString + dr("TP_NAME").ToString + Environment.NewLine
                    End If
                End If
                If Not IsDBNull(dr("TP_CURR_ADD_LINE1")) Then
                    If Not String.IsNullOrEmpty(dr("TP_CURR_ADD_LINE1").ToString) Then
                        strTempString = strTempString + dr("TP_CURR_ADD_LINE1").ToString
                    End If
                End If
                If Not IsDBNull(dr("TP_CURR_ADD_LINE2")) Then
                    If Not String.IsNullOrEmpty(dr("TP_CURR_ADD_LINE2").ToString) Then
                        If Right(Trim(dr("TP_CURR_ADD_LINE2")), 1) = "," Then
                            strTempString = strTempString + " " + dr("TP_CURR_ADD_LINE2").ToString
                        Else
                            strTempString = strTempString + ", " + dr("TP_CURR_ADD_LINE2").ToString
                        End If
                    End If
                End If
                If Not IsDBNull(dr("TP_CURR_ADD_LINE3")) Then
                    If Not String.IsNullOrEmpty(dr("TP_CURR_ADD_LINE3").ToString) Then
                        If Right(Trim(dr("TP_CURR_ADD_LINE3")), 1) = "," Then
                            strTempString = strTempString + " " + dr("TP_CURR_ADD_LINE3").ToString
                        Else
                            strTempString = strTempString + ", " + dr("TP_CURR_ADD_LINE3").ToString
                        End If
                    End If
                End If
                If Not IsDBNull(dr("TP_CURR_POSTCODE")) Then
                    If Not String.IsNullOrEmpty(dr("TP_CURR_POSTCODE").ToString) Then
                        strTempString = strTempString + Environment.NewLine + dr("TP_CURR_POSTCODE").ToString
                    End If
                End If
                If Not IsDBNull(dr("TP_CURR_CITY")) Then
                    If Not String.IsNullOrEmpty(dr("TP_CURR_CITY").ToString) Then
                        strTempString = strTempString + " " + dr("TP_CURR_CITY").ToString + Environment.NewLine
                    End If
                End If
                If Not IsDBNull(dr("TP_CURR_STATE")) Then
                    If Not String.IsNullOrEmpty(dr("TP_CURR_STATE").ToString) Then
                        strTempString = strTempString + dr("TP_CURR_STATE").ToString
                    End If
                End If
                If Not String.IsNullOrEmpty(strTempString) Then
                    pdfFormFields.SetField(pdfFieldPath & "Slip3", strTempString.ToString.ToUpper)
                End If

                strTempString = ""
                If Not IsDBNull(dr("TP_IC_NEW_1")) Then
                    If Not String.IsNullOrEmpty(dr("TP_IC_NEW_1").ToString) Then
                        strTempString = strTempString + dr("TP_IC_NEW_1").ToString
                    End If
                End If
                If Not IsDBNull(dr("TP_IC_NEW_2")) Then
                    If Not String.IsNullOrEmpty(dr("TP_IC_NEW_2").ToString) Then
                        strTempString = strTempString + dr("TP_IC_NEW_2").ToString
                    End If
                End If
                If Not IsDBNull(dr("TP_IC_NEW_3")) Then
                    If Not String.IsNullOrEmpty(dr("TP_IC_NEW_3").ToString) Then
                        strTempString = strTempString + dr("TP_IC_NEW_3").ToString
                    End If
                End If
                If Not String.IsNullOrEmpty(strTempString) Then
                    pdfFormFields.SetField(pdfFieldPath & "Slip4", strTempString)
                End If
            End If
            dr.Close()

            dr = datHandler.GetDataReader("Select TC_BALANCE_TAX_PAYABLE from tax_computation where tc_ref_no= '" & pdfForm.GetRefNo & "' and tc_ya= '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                If Not IsDBNull("TC_BALANCE_TAX_PAYABLE") Then
                    pdfFormFields.SetField(pdfFieldPath & "Slip2", FormatFloatingAmount(dr("TC_BALANCE_TAX_PAYABLE").ToString, True))
                End If
            End If
            dr.Close()

            pdfFormFields.SetField(pdfFieldPath & "Slip5", "")
            pdfFormFields.SetField(pdfFieldPath & "Slip6", "")
            pdfFormFields.SetField(pdfFieldPath & "Slip7", "")
            pdfFormFields.SetField(pdfFieldPath & "Slip8", "")
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
            strTemp = Format(dtTemp, "ddMMyyyy").ToString.Replace("-", "").Replace("/", "")
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

#End Region
End Class
