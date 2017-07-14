Imports iTextSharp.text.pdf
Imports System.Data.SqlClient
Imports System.IO
Imports System.Math

Public Class clsBorangM2016
    Private Const pdfSubFormName = "topmostSubform[0]."
    Dim pdfForm As New clsPDFMaker
    Dim pdfFormFields As AcroFields
    Dim datHandler As New clsDataHandler("")
    Dim RefName As String, PnLMainBusiness As String
    Dim strSQL As String = Nothing
#Region "CStor"

    Public Sub New()

        datHandler = New clsDataHandler(pdfForm.GetFormType)
        pdfFormFields = pdfForm.GetStamper.AcroFields
        CheckFieldEmpty()
        'Call your page number here()
        Page1()
        Page2()
        Page3()
        'Page4()
        'Page5()
        'Page6()
        'Page7()
        'Page8()
        'Page9()
        'Page10()
        'Page11()
        'Page12()

        pdfForm.OpenFile()
        pdfForm.CloseStamper()
    End Sub

#End Region

#Region "Insert the page function here"

    Private Sub Page1()
        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim prmOledb(0) As SqlParameter
        Dim strArray(1) As String
        Dim strLine As String = ""

        Dim dr2 As SqlDataReader = Nothing
        Dim dblTotalValue As Double = 0
        Dim nTotal As Double = 0
        Dim grossOtherRate As Double = 0

        Dim ComputationKey As String

        Try
            'Master Data 
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page3[0]."


            ds = datHandler.GetData("SELECT TP_NAME, TP_REF_NO_PREFIX, (TP_REF_NO1 + TP_REF_NO2 + TP_REF_NO3)," _
                        & " (TP_IC_NEW_1 + TP_IC_NEW_2 + TP_IC_NEW_3), TP_IC_OLD, TP_POLICE_NO, TP_ARMY_NO," _
                        & " TP_PASSPORT_NO, TP_PASSWPORTDUEDATE, TP_RESIDENCE, TP_COUNTRY, TP_GENDER, TP_STATUS," _
                        & " TP_DATE_MARRIAGE, TP_DATE_DIVORCE, TP_TYPE_ASSESSMENT, TP_KUP," _
                        & " (TP_CURR_ADD_LINE1 + ', ' + TP_CURR_ADD_LINE2 + ', ' + TP_CURR_ADD_LINE3), TP_CURR_POSTCODE," _
                        & " TP_CURR_CITY, TP_CURR_STATE, TP_ASSESSMENTON" _
                        & " FROM TAXP_PROFILE WHERE (TP_REF_NO1 + TP_REF_NO2 + TP_REF_NO3)=@ref_no", prmOledb)

            'strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 28)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "I[0]", ds.Tables(0).Rows(0).Item(0).ToString().ToUpper)
            'pdfFormFields.SetField(pdfFieldPath & "I_2[0]", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "II[0]", ds.Tables(0).Rows(0).Item(1).ToString + " " + ds.Tables(0).Rows(0).Item(2).ToString)
            'pdfFormFields.SetField(pdfFieldPath & "II_2[0]", ds.Tables(0).Rows(0).Item(2).ToString)
            pdfFormFields.SetField(pdfFieldPath & "III[0]", FormatICNumber(ds.Tables(0).Rows(0).Item(3).ToString))
            pdfFormFields.SetField(pdfFieldPath & "IV[0]", ds.Tables(0).Rows(0).Item(4).ToString)
            pdfFormFields.SetField(pdfFieldPath & "V[0]", ds.Tables(0).Rows(0).Item(5).ToString)
            pdfFormFields.SetField(pdfFieldPath & "VI[0]", ds.Tables(0).Rows(0).Item(6).ToString)
            pdfFormFields.SetField(pdfFieldPath & "VII[0]", ds.Tables(0).Rows(0).Item(7).ToString)
            If Not IsDBNull(ds.Tables(0).Rows(0).Item(8)) Then
                If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(8).ToString) Then
                    pdfFormFields.SetField(pdfFieldPath & "VIII[0]", FormatDateSlash(ds.Tables(0).Rows(0).Item(8)))
                End If
            End If

            'Declaration
            If pdfForm.GetDeclarationReturn = 1 Then
                dr = datHandler.GetDataReader("SELECT * FROM [TAXP_PROFILE] WHERE [TP_REF_NO1] & [TP_REF_NO2] & [TP_REF_NO3] = '" & pdfForm.GetRefNo & "'")
                If dr.Read() Then
                    'strArray = SplitText(RefName, 28)
                    pdfFormFields.SetField(pdfFieldPath & "Akuan1[0]", RefName)
                    'pdfFormFields.SetField(pdfFieldPath & "Akuan1_2[0]", strArray(1))
                    If Len(Trim(dr("TP_IC_NEW_1") + Trim(dr("TP_IC_NEW_2")) + Trim(dr("TP_IC_NEW_3")))) > 0 Then
                        strLine = Trim(dr("TP_IC_NEW_1")) + "-" + Trim(dr("TP_IC_NEW_2")) + "-" + Trim(dr("TP_IC_NEW_3"))
                    ElseIf Len(Trim(dr("TP_PASSPORT_NO"))) > 0 Then
                        strLine = (dr("TP_PASSPORT_NO"))
                    ElseIf Len(Trim(dr("TP_POLICE_NO"))) > 0 Then
                        strLine = (dr("TP_POLICE_NO"))
                    ElseIf Len(Trim(dr("TP_ARMY_NO"))) > 0 Then
                        strLine = (dr("TP_ARMY_NO"))
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "Akuan2", strLine)
                    pdfFormFields.SetField(pdfFieldPath & "Akuan3", pdfForm.GetDeclarationReturn)
                    pdfFormFields.SetField(pdfFieldPath & "Akuan4", "")
                    pdfFormFields.SetField(pdfFieldPath & "Akuan5", FormatDateDash(pdfForm.GetDeclarationDate))
                End If
                dr.Close()
            ElseIf pdfForm.GetDeclarationReturn = 2 Then
                strArray = SplitText(pdfForm.GetDeclarationBy.ToString.ToUpper, 28)
                pdfFormFields.SetField(pdfFieldPath & "Akuan1_1[0]", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "Akuan1_2[0]", strArray(1))
                pdfFormFields.SetField(pdfFieldPath & "Akuan2", pdfForm.GetDeclarationID)
                pdfFormFields.SetField(pdfFieldPath & "Akuan3", pdfForm.GetDeclarationReturn)
                pdfFormFields.SetField(pdfFieldPath & "Akuan4", RefName)
                pdfFormFields.SetField(pdfFieldPath & "Akuan5", FormatDateDash(pdfForm.GetDeclarationDate))
            End If

            'Part A
            pdfFormFields.SetField(pdfFieldPath & "A1[0]", ds.Tables(0).Rows(0).Item(10).ToString)
            'pdfFormFields.SetField(pdfFieldPath & "A2[0]", ds.Tables(0).Rows(0).Item(10).ToString)
            pdfFormFields.SetField(pdfFieldPath & "A3[0]", ds.Tables(0).Rows(0).Item(11).ToString)
            pdfFormFields.SetField(pdfFieldPath & "A4[0]", ds.Tables(0).Rows(0).Item(12).ToString)
            If Trim(ds.Tables(0).Rows(0).Item(12).ToString) = "2" Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(13)) Or ds.Tables(0).Rows(0).Item(13).ToString <> "" Then
                    pdfFormFields.SetField(pdfFieldPath & "A5[0]", FormatDateSlash(ds.Tables(0).Rows(0).Item(13)))
                End If
            ElseIf Trim(ds.Tables(0).Rows(0).Item(12).ToString) = "3" Or Trim(ds.Tables(0).Rows(0).Item(12)) = "4" Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(14)) Or ds.Tables(0).Rows(0).Item(14).ToString <> "" Then
                    pdfFormFields.SetField(pdfFieldPath & "A5[0]", FormatDateSlash(ds.Tables(0).Rows(0).Item(14)))
                End If
            Else
                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "")
            End If
            If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(15).ToString) Then
                If ds.Tables(0).Rows(0).Item(15).ToString = "1" Then
                    If ds.Tables(0).Rows(0).Item(21).ToString = "1" Then
                        pdfFormFields.SetField(pdfFieldPath & "A6[0]", "1")
                    ElseIf ds.Tables(0).Rows(0).Item(21).ToString = "2" Then
                        pdfFormFields.SetField(pdfFieldPath & "A6[0]", "2")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A6[0]", "")
                    End If
                ElseIf ds.Tables(0).Rows(0).Item(15).ToString = "2" Then
                    pdfFormFields.SetField(pdfFieldPath & "A6[0]", "3")
                ElseIf ds.Tables(0).Rows(0).Item(15).ToString = "3" Then
                    If ds.Tables(0).Rows(0).Item(12).ToString = "2" Then
                        pdfFormFields.SetField(pdfFieldPath & "A6[0]", "4")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A6[0]", "5")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "A6[0]", "")
                End If
            End If

            'lyeyc
            If ds.Tables(0).Rows(0).Item(16).ToString = "1" Then
                pdfFormFields.SetField(pdfFieldPath & "A7[0]", "1")
            Else
                pdfFormFields.SetField(pdfFieldPath & "A7[0]", "2")
            End If
            'lyeyc (end)

            If pdfForm.GetRecordKeep = 1 Then
                pdfFormFields.SetField(pdfFieldPath & "A8[0]", "1")
            Else
                pdfFormFields.SetField(pdfFieldPath & "A8[0]", "2")
            End If

            ReDim strArray(2)
            strArray = SplitText(ds.Tables(0).Rows(0).Item(17).ToString().Replace(",,", ",").Replace(", ,", ",").ToUpper, 26)
            pdfFormFields.SetField(pdfFieldPath & "A9_1[0]", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "A9_2[0]", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "A9_3[0]", strArray(2))
            pdfFormFields.SetField(pdfFieldPath & "A9_4[0]", ds.Tables(0).Rows(0).Item(18).ToString)
            pdfFormFields.SetField(pdfFieldPath & "A9_5[0]", ds.Tables(0).Rows(0).Item(19).ToString.ToUpper)
            pdfFormFields.SetField(pdfFieldPath & "A9_6[0]", ds.Tables(0).Rows(0).Item(20).ToString.ToUpper)

            'Master Data
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            ds = datHandler.GetData("SELECT TP_LAST_PASSPORT_NO, TP_DOB, TP_WORKER_APPROVEDATE, TP_COM_ADD_STATUS FROM TAXP_PROFILE2 WHERE" _
                        & " TP_REF_NO= ?", prmOledb)
            pdfFormFields.SetField(pdfFieldPath & "VIIa[0]", ds.Tables(0).Rows(0).Item(0).ToString)
            'lyeyc
            If Not IsDBNull(ds.Tables(0).Rows(0).Item(3)) Then
                If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(3).ToString) Then
                    If ds.Tables(0).Rows(0).Item(3).ToString = "1" Then
                        pdfFormFields.SetField(pdfFieldPath & "A9_7", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A9_7", "")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "A9_7", "")
                End If
            Else
                pdfFormFields.SetField(pdfFieldPath & "A9_7", "")
            End If
            'lyeyc (end)
            If Not IsDBNull(ds.Tables(0).Rows(0).Item(1)) Then
                If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(1).ToString) Then
                    pdfFormFields.SetField(pdfFieldPath & "IX[0]", FormatDateSlash(ds.Tables(0).Rows(0).Item(1)))
                End If
            End If
            If Not IsDBNull(ds.Tables(0).Rows(0).Item(2)) Then
                pdfFormFields.SetField(pdfFieldPath & "A9[0]", "1")
                pdfFormFields.SetField(pdfFieldPath & "A9a[0]", FormatDate(ds.Tables(0).Rows(0).Item(2)))
                ' pdfFormFields.SetField(pdfFieldPath & "A4[0]", FormatDate(ds.Tables(0).Rows(0).Item(11)))
            Else
                pdfFormFields.SetField(pdfFieldPath & "A9[0]", "2")
            End If
            'Initialise
            pdfFormFields.SetField(pdfFieldPath & "IXI_1[0]", "")
            pdfFormFields.SetField(pdfFieldPath & "IXI_2[0]", "")
            pdfFormFields.SetField(pdfFieldPath & "IXI_3[0]", "")
            pdfFormFields.SetField(pdfFieldPath & "IXI_4[0]", "")

            Select Case (GetStatusOfTax())
                Case "REPAYABLE"
                    pdfFormFields.SetField(pdfFieldPath & "IXI_1[0]", "X")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_2[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_3[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_4[0]", "")
                Case "EXCESS"
                    pdfFormFields.SetField(pdfFieldPath & "IXI_1[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_2[0]", "X")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_3[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_4[0]", "")
                Case "BALANCE"
                    pdfFormFields.SetField(pdfFieldPath & "IXI_1[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_2[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_3[0]", "X")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_4[0]", "")
                Case "NIL"
                    pdfFormFields.SetField(pdfFieldPath & "IXI_1[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_2[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_3[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "IXI_4[0]", "X")
            End Select

            ReDim prmOledb(1)
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            prmOledb(1) = New SqlParameter("@ya", pdfForm.GetYA)
            ds = datHandler.GetData("SELECT SUM(cast(TCA_CBL as money)) FROM TAX_ADJUSTED_LOSS WHERE TC_KEY IN (SELECT TC_KEY FROM TAX_COMPUTATION WHERE TC_REF_NO=@ref_no AND TC_YA=@ya)", prmOledb)

            If (ds.Tables(0).Rows.Count > 0) Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) Then
                    If (CDbl(ds.Tables(0).Rows(0).Item(0).ToString) > 0) Then
                        pdfFormFields.SetField(pdfFieldPath & "A8a[0]", "1")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A8a[0]", "2")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "A8a[0]", "2")
                End If
            Else
                pdfFormFields.SetField(pdfFieldPath & "A8a[0]", "2")
            End If

            ds.Dispose()

            'C7 - C20
            dr = datHandler.GetDataReader("Select TC_BUSINESS, TC_PARTNERSHIP, TC_STATUTORY_INCOME, TC_BUSINESSLOSS_BF, TC_AGGREGATE_BUS_INCOME," _
                                & " TC_EMPLOYMENT_INCOME, TC_DIVIDEND, (cdbl(TC_INTEREST) + cdbl(TC_DISCOUNT)), " _
                                & " (cdbl(TC_RENTAL_ROYALTY)+cdbl(TC_PREMIUM)), TC_PENSION_AND_ETC," _
                                & " (cdbl(TC_OTHER_GAIN_PROFIT) + cdbl(TC_SEC4A)), TC_ADDITION_43," _
                                & " TC_AGGREGATE_OTHER_SRC, TC_AGGREGATE_INCOME, TC_BUSINESSLOSS_CY," _
                                & " TC_TOTAL1, TC_EXEMPT_CLAIM, TC_EXEMPT_COUNTRY, TC_EXEMPT_AMOUNT, (cdbl(TC_INTEREST)+cdbl(TC_DISCOUNT)+cdbl(TC_RENTAL_ROYALTY)+cdbl(TC_PREMIUM)+cdbl(TC_PENSION_AND_ETC)+cdbl(TC_OTHER_GAIN_PROFIT)+cdbl(TC_SEC4A)) from tax_computation where" _
                                & " tc_ref_no ='" & pdfForm.GetRefNo & "' and tc_ya ='" & pdfForm.GetYA & "'")
            'weihong TC_EXEMPT_AMOUNT
            If dr.Read() Then
                If Not IsDBNull(dr("TC_BUSINESS")) Then
                    pdfFormFields.SetField(pdfFieldPath & "B1[0]", CDbl(FormatNumber((dr("TC_BUSINESS")), 0)))
                End If
                If Not IsDBNull(dr("TC_PARTNERSHIP")) Then
                    pdfFormFields.SetField(pdfFieldPath & "B2[0]", CDbl(FormatNumber((dr("TC_PARTNERSHIP")), 0)))
                End If

                If Not IsDBNull(dr("TC_STATUTORY_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C7[0]", CDbl(FormatNumber((dr("TC_STATUTORY_INCOME")), 0)))
                End If
                If Not IsDBNull(dr("TC_BUSINESSLOSS_BF")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C8[0]", CDbl(FormatNumber((dr("TC_BUSINESSLOSS_BF")), 0)))
                End If
                If Not IsDBNull(dr("TC_AGGREGATE_BUS_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C9[0]", CDbl(FormatNumber((dr("TC_AGGREGATE_BUS_INCOME")), 0)))
                End If
                If CDbl(dr("TC_EXEMPT_CLAIM")) > 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "C10_1[0]", dr("TC_EXEMPT_CLAIM"))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "C10_1[0]", "")
                End If
                If Not IsDBNull(dr("TC_EXEMPT_COUNTRY")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C10_2[0]", dr("TC_EXEMPT_COUNTRY").ToString)
                End If
                If Not IsDBNull(dr("TC_EMPLOYMENT_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C10_3[0]", CDbl(FormatNumber((dr("TC_EMPLOYMENT_INCOME")), 0)))
                End If

                dr2 = datHandler.GetDataReader("Select OS_DV_STAT_INCOME, OS_RT_RENTAL_BF" _
                    & " from income_othersource where" _
                    & " os_ref_no ='" & pdfForm.GetRefNo & "' and os_ya ='" & pdfForm.GetYA & "'")
                If dr2.Read() Then
                    If Not (IsDBNull(dr2.Item(0)) And IsDBNull(dr2.Item(1))) Then
                        pdfFormFields.SetField(pdfFieldPath & "B7[0]", Fix((CDbl(dr2.Item(0))) + CDbl(dr2.Item(1))))
                    End If

                    '--- B3 ---
                    If Not IsDBNull(dr.Item(19)) Then
                        pdfFormFields.SetField(pdfFieldPath & "B8[0]", (CDbl(dr.Item(19)) - CDbl(dr2.Item(1))))
                    End If
                End If
                dr2.Close()

                'weihong
                If Not IsDBNull(dr("TC_EXEMPT_AMOUNT")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C10_4[0]", CDbl(FormatNumber((dr("TC_EXEMPT_AMOUNT")), 0)))
                End If
                'weihong
                If Not IsDBNull(dr("TC_DIVIDEND")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C11[0]", CDbl(FormatNumber((dr("TC_DIVIDEND")), 0)))
                End If
                If Not IsDBNull(dr.Item(5)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C12[0]", CDbl(FormatNumber((dr.Item(5).ToString), 0)))
                End If
                If Not IsDBNull(dr.Item(6)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C13[0]", CDbl(FormatNumber((dr.Item(6).ToString), 0)))
                End If
                If Not IsDBNull(dr("TC_PENSION_AND_ETC")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C14[0]", CDbl(FormatNumber((dr("TC_PENSION_AND_ETC")), 0)))
                End If
                If Not IsDBNull(dr.Item(8)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C15[0]", CDbl(FormatNumber((dr.Item(8).ToString), 0)))
                End If
                If Not IsDBNull(dr("TC_ADDITION_43")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C16[0]", CDbl(FormatNumber((dr("TC_ADDITION_43")), 0)))
                End If
                If Not IsDBNull(dr("TC_AGGREGATE_OTHER_SRC")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C17[0]", CDbl(FormatNumber((dr("TC_AGGREGATE_OTHER_SRC")), 0)))
                End If
                If Not IsDBNull(dr("TC_AGGREGATE_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C18[0]", CDbl(FormatNumber((dr("TC_AGGREGATE_INCOME")), 0)))
                End If
                If Not IsDBNull(dr("TC_BUSINESSLOSS_CY")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C19[0]", CDbl(FormatNumber((dr("TC_BUSINESSLOSS_CY")), 0)))
                End If
                If Not IsDBNull(dr("TC_TOTAL1")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C20[0]", CDbl(FormatNumber((dr("TC_TOTAL1")), 0)))
                End If
            Else

            End If
            dr.Close()



            'Part C
            dr = datHandler.GetDataReader("select TC_KEY, TC_AGGREGATE_INCOME, TC_PROSPECTING, TC_QUALIFYING_AG_EXP, TC_DONATION_GIFT, TC_TOTAL2," _
                                    & " TC_4, TC_3, TC_TOTAL_INCOME_1, TC_TOTAL_INCOME_2, TC_INCOME_TRANSFER_FROM_HW, TC_TOTAL_INCOME_3" _
                                    & " from tax_computation where" _
                                    & " tc_ref_no ='" & pdfForm.GetRefNo & "' and tc_ya ='" & pdfForm.GetYA & "'")
            If dr.Read Then

                If Not IsDBNull(dr("TC_PROSPECTING")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C21[0]", CDbl(FormatNumber((dr("TC_PROSPECTING")), 0)))
                End If

                If Not IsDBNull(dr("TC_DONATION_GIFT")) Then
                    pdfFormFields.SetField(pdfFieldPath & "B12[0]", CDbl(FormatNumber((dr("TC_DONATION_GIFT")), 0)))
                End If

                If Not IsDBNull(dr("TC_TOTAL_INCOME_1")) Then
                    pdfFormFields.SetField(pdfFieldPath & "B13[0]", CDbl(FormatNumber((dr("TC_TOTAL_INCOME_1")), 0)))
                End If

                If Not IsDBNull(dr("TC_TOTAL2")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C22[0]", CDbl(FormatNumber((dr("TC_TOTAL2")), 0)))
                End If

                'C23 - 'C30
                dr2 = datHandler.GetDataReader("select TCG_KEY, TCG_AMOUNT" _
                                        & " from tax_gifts where" _
                                        & " tc_key =" & dr("TC_KEY"))
                Do While dr2.Read()
                    If Not IsDBNull(dr2("TCG_KEY")) Then

                        Select Case dr2("TCG_KEY")
                            Case "9"
                                pdfFormFields.SetField(pdfFieldPath & "C23[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                            Case "1"
                                pdfFormFields.SetField(pdfFieldPath & "C23A[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "7"
                                pdfFormFields.SetField(pdfFieldPath & "C24[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "8"
                                pdfFormFields.SetField(pdfFieldPath & "C25_1[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "2"
                                pdfFormFields.SetField(pdfFieldPath & "C26[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                            Case "3"
                                pdfFormFields.SetField(pdfFieldPath & "C27[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                            Case "4"
                                pdfFormFields.SetField(pdfFieldPath & "C28[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                            Case "5"
                                pdfFormFields.SetField(pdfFieldPath & "C29[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                            Case "6"
                                pdfFormFields.SetField(pdfFieldPath & "C30[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))

                        End Select
                    End If
                Loop
                dr2.Close()

                'C26 Total restrict to 7% of C28
                If Not IsDBNull(dr("TC_AGGREGATE_INCOME")) Then
                    If Not String.IsNullOrEmpty(dr("TC_AGGREGATE_INCOME").ToString) Then
                        If dblTotalValue >= CDbl(dr("TC_AGGREGATE_INCOME")) * 0.07 Then
                            pdfFormFields.SetField(pdfFieldPath & "C25_2[0]", CDbl(FormatNumber(((CDbl(dr("TC_AGGREGATE_INCOME")) * 0.07).ToString), 0)))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "C25_2[0]", CDbl(FormatNumber((dblTotalValue.ToString), 0)))
                        End If
                    End If
                End If

                'C31 - C35_2
                'NGOHCS B2010.2
                dr2 = datHandler.GetDataReader("select SUM(cast(TCG_AMOUNT as money))" _
                                       & " from tax_gifts where" _
                                       & " tc_key =" & dr("TC_KEY"))
                If dr2.Read() Then
                    If Not IsDBNull(dr.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "C31[0]", CDbl(dr("TC_TOTAL2")) - CDbl(dr2.Item(0)))
                        nTotal = nTotal + (CDbl(dr("TC_TOTAL2")) - CDbl(dr2.Item(0)))
                    End If
                End If
                dr2.Close()
                'NGOHCS B2010.2 END
                'If Not IsDBNull(dr("TC_4")) Then
                '    pdfFormFields.SetField(pdfFieldPath & "C31[0]", CDbl(FormatNumber((dr("TC_4")), 0)))
                '    nTotal = CDbl(dr("TC_4"))
                'End If

                If Not IsDBNull(dr("TC_3")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C32[0]", CDbl(FormatNumber((dr("TC_3")), 0)))
                    nTotal = nTotal + CDbl(dr("TC_3"))
                End If
                dr2 = datHandler.GetDataReader("select INTEREST, ROYALTIES, SECTION4A, OTHERINCOME" _
                                                        & " from CHARGEABLE_INCOME where" _
                                                        & " tc_key =" & dr("TC_KEY"))
                If dr2.Read Then
                    If Not IsDBNull(dr2("INTEREST")) Then
                        grossOtherRate = grossOtherRate + dr2("INTEREST")
                        'pdfFormFields.SetField(pdfFieldPath & "C33a[0]", CDbl(FormatNumber((dr2("INTEREST")), 0)))
                        'nTotal = nTotal + CDbl(dr2("INTEREST"))
                    Else
                        'pdfFormFields.SetField(pdfFieldPath & "C33a[0]", 0)
                    End If
                    If Not IsDBNull(dr2("ROYALTIES")) Then
                        grossOtherRate = grossOtherRate + dr2("ROYALTIES")
                        'pdfFormFields.SetField(pdfFieldPath & "C33b[0]", CDbl(FormatNumber((dr2("ROYALTIES")), 0)))
                        'nTotal = nTotal + CDbl(dr2("ROYALTIES"))
                    Else
                        'pdfFormFields.SetField(pdfFieldPath & "C33b[0]", 0)
                    End If
                    If Not IsDBNull(dr2("SECTION4A")) Then
                        grossOtherRate = grossOtherRate + dr2("SECTION4A")
                        'pdfFormFields.SetField(pdfFieldPath & "C33c[0]", CDbl(FormatNumber((dr2("SECTION4A")), 0)))
                        'nTotal = nTotal + CDbl(dr2("SECTION4A"))
                    Else
                        'pdfFormFields.SetField(pdfFieldPath & "C33c[0]", 0)
                    End If
                    If Not IsDBNull(dr2("OTHERINCOME")) Then
                        grossOtherRate = grossOtherRate + dr2("OTHERINCOME")
                        'pdfFormFields.SetField(pdfFieldPath & "C33d_1[0]", "")
                        'pdfFormFields.SetField(pdfFieldPath & "C33d_2[0]", CDbl(FormatNumber((dr2("OTHERINCOME")), 0)))
                        'nTotal = nTotal + CDbl(dr2("OTHERINCOME"))
                    Else
                        'pdfFormFields.SetField(pdfFieldPath & "C33d_1[0]", "")
                        'pdfFormFields.SetField(pdfFieldPath & "C33d_2[0]", 0)
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "B15_2[0]", grossOtherRate)
                Else
                    'pdfFormFields.SetField(pdfFieldPath & "C33a[0]", 0)
                    'pdfFormFields.SetField(pdfFieldPath & "C33b[0]", 0)
                    'pdfFormFields.SetField(pdfFieldPath & "C33c[0]", 0)
                    'pdfFormFields.SetField(pdfFieldPath & "C33d_1[0]", "")
                    'pdfFormFields.SetField(pdfFieldPath & "C33d_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "B15_2[0]", 0)
                End If
                dr2.Close()
                pdfFormFields.SetField(pdfFieldPath & "C34[0]", nTotal)

                If Not IsDBNull(dr("TC_INCOME_TRANSFER_FROM_HW")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C35_2[0]", CDbl(FormatNumber((dr("TC_INCOME_TRANSFER_FROM_HW")), 0)))
                End If

                'C35_1
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
                    If Not IsDBNull(dr2("TP_HW_TYPEOFINCOME")) And Not String.IsNullOrEmpty(dr2("TP_HW_TYPEOFINCOME").ToString) Then
                        If dr2("TP_TYPE_ASSESSMENT") = "1" Then
                            If (dr2("TP_GENDER") = "1" And dr2("TP_ASSESSMENTON") = "1") Or _
                                (dr2("TP_GENDER") = "2" And dr2("TP_ASSESSMENTON") = "2") Then
                                If dr2("TP_HW_TYPEOFINCOME").ToString = "1" Or boolWithBusiness = True Then
                                    pdfFormFields.SetField(pdfFieldPath & "C35_1[0]", "1")
                                Else
                                    pdfFormFields.SetField(pdfFieldPath & "C35_1[0]", "2")
                                End If
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "C35_1[0]", "")
                            End If
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "C35_1[0]", "")
                        End If
                    End If

                    'C36
                    'NGOHCS B2010.2
                    'pdfFormFields.SetField(pdfFieldPath & "C36[0]", nTotal + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW")))
                    If Not IsDBNull(dr2("TP_STATUS")) Then
                        If dr2("TP_STATUS") = "1" Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        ElseIf dr2("TP_STATUS") = "2" And CDbl(pdfForm.GetYA) >= 2007 And dr2("TP_TYPE_ASSESSMENT") = "3" Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        ElseIf dr2("TP_STATUS") = "2" And CDbl(pdfForm.GetYA) = 2006 And dr2("TP_TYPE_ASSESSMENT") = "1" And CDbl(dr("TC_INCOME_TRANSFER_FROM_HW")) = 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        ElseIf (dr2("TP_GENDER") = "1" And dr2("TP_STATUS") = "2" And dr2("TP_TYPE_ASSESSMENT") = "1" And _
                            dr2("TP_ASSESSMENTON") = "1") Or (dr2("TP_GENDER") = "2" And dr2("TP_STATUS") = "2" And dr2("TP_TYPE_ASSESSMENT") = "1" And dr2("TP_ASSESSMENTON") = "2") Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", CDbl(FormatNumber(((CDbl(dr("TC_TOTAL_INCOME_2")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))).ToString), 0)))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        End If
                    End If

                End If
                dr2.Close()
            End If
            dr.Close()

            dr2 = datHandler.GetDataReader("select * from taxp_profile2 where tp_ref_no= '" & pdfForm.GetRefNo & "'")
            If dr2.Read() Then

                If Not IsDBNull(dr2("TP_BUSINESS_ECOMMERCE")) AndAlso dr2("TP_BUSINESS_ECOMMERCE") = "1" Then
                    pdfFormFields.SetField(pdfFieldPath & "D7A_12[0]", "1")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "D7A_12[0]", "2")
                End If

                If Not IsDBNull(dr2("TP_BWA")) Then
                    pdfFormFields.SetField(pdfFieldPath & "D7A_13[0]", dr2("TP_BWA"))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "D7A_13[0]", "")
                End If

                If Not IsDBNull(dr2("TP_JKDM")) AndAlso dr2("TP_JKDM") = "1" Then
                    pdfFormFields.SetField(pdfFieldPath & "D7A_14[0]", "1")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "D7A_14[0]", "2")
                End If

                If Not IsDBNull(dr2("TP_GST")) Then
                    pdfFormFields.SetField(pdfFieldPath & "D7A_17[0]", dr2("TP_GST"))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "D7A_17[0]", "")
                End If

                If Not IsDBNull(dr2("TP_DISPOSAL1976")) AndAlso dr2("TP_DISPOSAL1976") = "1" Then
                    pdfFormFields.SetField(pdfFieldPath & "D7A_15[0]", "1")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "D7A_15[0]", "2")
                End If

                If Not IsDBNull(dr2("TP_LDMN_DISPOSAL")) AndAlso dr2("TP_LDMN_DISPOSAL") = "1" Then
                    pdfFormFields.SetField(pdfFieldPath & "D7A_16[0]", "1")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "D7A_16[0]", "2")
                End If

            End If
            dr2.Close()
            'Part D
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            prmOledb(1) = New SqlParameter("@YA", pdfForm.GetYA)
            'weihong
            ds = datHandler.GetData("SELECT TC_CHARGEABLE_INCOME, TC_TOTAL_INCOME_TAX, TC_INCOME_TAX_CHARGED," _
                            & " TC_SEC110_DIVIDEND, TC_SEC110_OTHERS, TC_SEC130, TC_5, TC_TAX_PAYABLE, TC_TAX_REPAYMENT, TC_KEY, TC_2, TC_TAX_SCH1_INCOME, TC_TAX_SCH1_TAX" _
                            & " FROM TAX_COMPUTATION WHERE TC_REF_NO=@ref_no AND TC_YA=@ya", prmOledb(0), prmOledb(1))
            pdfFormFields.SetField(pdfFieldPath & "D1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(0).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "D3_1[0]", Left((FormatFloatingAmount(ds.Tables(0).Rows(0).Item(1).ToString, True)), (Len(FormatFloatingAmount(ds.Tables(0).Rows(0).Item(1).ToString, True)) - 2)))
            pdfFormFields.SetField(pdfFieldPath & "D3_2[0]", Right((FormatFloatingAmount(FormatNumber(ds.Tables(0).Rows(0).Item(1), 2).ToString, True)), 2))
            pdfFormFields.SetField(pdfFieldPath & "D5[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(2).ToString, 2).ToString.Replace(".", "")))

            pdfFormFields.SetField(pdfFieldPath & "D6_1[0]", Left((FormatFloatingAmount(ds.Tables(0).Rows(0).Item(3).ToString, True)), (Len(FormatFloatingAmount(ds.Tables(0).Rows(0).Item(3).ToString, True)) - 2)))
            pdfFormFields.SetField(pdfFieldPath & "D6_2[0]", Right((FormatFloatingAmount(FormatNumber(ds.Tables(0).Rows(0).Item(3), 2).ToString, True)), 2))
            pdfFormFields.SetField(pdfFieldPath & "D7_1[0]", Left((FormatFloatingAmount(ds.Tables(0).Rows(0).Item(4).ToString, True)), (Len(FormatFloatingAmount(ds.Tables(0).Rows(0).Item(4).ToString, True)) - 2)))
            pdfFormFields.SetField(pdfFieldPath & "D7_2[0]", Right((FormatFloatingAmount(FormatNumber(ds.Tables(0).Rows(0).Item(4), 2).ToString, True)), 2))
            pdfFormFields.SetField(pdfFieldPath & "D8_1[0]", Left((FormatFloatingAmount(ds.Tables(0).Rows(0).Item(10).ToString, True)), (Len(FormatFloatingAmount(ds.Tables(0).Rows(0).Item(10).ToString, True)) - 2)))
            pdfFormFields.SetField(pdfFieldPath & "D8_2[0]", Right((FormatFloatingAmount(FormatNumber(ds.Tables(0).Rows(0).Item(10), 2).ToString, True)), 2))
            pdfFormFields.SetField(pdfFieldPath & "D9_1[0]", Left((FormatFloatingAmount(ds.Tables(0).Rows(0).Item(6).ToString, True)), (Len(FormatFloatingAmount(ds.Tables(0).Rows(0).Item(6).ToString, True)) - 2)))
            pdfFormFields.SetField(pdfFieldPath & "D9_2[0]", Right((FormatFloatingAmount(FormatNumber(ds.Tables(0).Rows(0).Item(6), 2).ToString, True)), 2))
            pdfFormFields.SetField(pdfFieldPath & "D10_1[0]", Left((FormatFloatingAmount(ds.Tables(0).Rows(0).Item(7).ToString, True)), (Len(FormatFloatingAmount(ds.Tables(0).Rows(0).Item(7).ToString, True)) - 2)))
            pdfFormFields.SetField(pdfFieldPath & "D10_2[0]", Right((FormatFloatingAmount(FormatNumber(ds.Tables(0).Rows(0).Item(7), 2).ToString, True)), 2))
            pdfFormFields.SetField(pdfFieldPath & "D11_1[0]", Left((FormatFloatingAmount(ds.Tables(0).Rows(0).Item(8).ToString, True)), (Len(FormatFloatingAmount(ds.Tables(0).Rows(0).Item(8).ToString, True)) - 2)))
            pdfFormFields.SetField(pdfFieldPath & "D11_2[0]", Right((FormatFloatingAmount(FormatNumber(ds.Tables(0).Rows(0).Item(8), 2).ToString, True)), 2))
            'weihong
            'If Not IsDBNull(ds.Tables(0).Rows(0).Item(11)) Then
            '    pdfFormFields.SetField(pdfFieldPath & "D2_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(11).ToString, 0).ToString.Replace(".", "")))
            'End If
            'If Not IsDBNull(ds.Tables(0).Rows(0).Item(12)) Then
            '    pdfFormFields.SetField(pdfFieldPath & "D2_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(12).ToString, 2).ToString.Replace(".", "")))
            'End If
            'pdfFormFields.SetField(pdfFieldPath & "D8A[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(10).ToString, 2).ToString.Replace(".", "")))

            'pdfFormFields.SetField(pdfFieldPath & "D9_2[0]", FormatFixedAmount(Right(CDbl(ds.Tables(0).Rows(0).Item(6)), 2)))

            'If (CDbl(ds.Tables(0).Rows(0).Item(6).ToString)) < CDbl(ds.Tables(0).Rows(0).Item(2).ToString) Then
            '    pdfFormFields.SetField(pdfFieldPath & "D10_1[0]", FormatFixedAmount(Left(CStr(CDbl(ds.Tables(0).Rows(0).Item(2) - CDbl(ds.Tables(0).Rows(0).Item(6)))), (Len(CStr(CDbl(ds.Tables(0).Rows(0).Item(2).ToString) - CDbl(ds.Tables(0).Rows(0).Item(6).ToString)) - 2)))))
            '    pdfFormFields.SetField(pdfFieldPath & "D10_2[0]", FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(2).ToString) - CDbl(ds.Tables(0).Rows(0).Item(6).ToString), 2).ToString.Replace(".", "").Replace(",", ""))
            '    pdfFormFields.SetField(pdfFieldPath & "D11[0]", FormatNumber(CDbl("0"), 2).ToString.Replace(".", ""))
            'Else
            '    pdfFormFields.SetField(pdfFieldPath & "D10[0]", FormatNumber(CDbl("0"), 2).ToString.Replace(".", ""))
            '    pdfFormFields.SetField(pdfFieldPath & "D11[0]", FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(6).ToString) - CDbl(ds.Tables(0).Rows(0).Item(2).ToString), 2).ToString.Replace(".", "").Replace(",", ""))
            'End If



            ComputationKey = ds.Tables(0).Rows(0).Item(9).ToString

            ds = datHandler.GetData("SELECT CHARGEABLE0, INCOME0, CHARGEABLE1, INCOME1, CHARGEABLE2, INCOME2," _
                            & " CHARGEABLE3, INCOME3, CHARGEABLE4, INCOME4, CHARGEABLE5, INCOME5, CHARGEABLE6, RATE1, INCOME6" _
                            & " FROM CHARGEABLE_INCOME WHERE TC_KEY=" + ComputationKey)
            'pdfFormFields.SetField(pdfFieldPath & "D2a_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(0).ToString, 0)))

            pdfFormFields.SetField(pdfFieldPath & "D2a_1[0]", Left((FormatFloatingAmount(ds.Tables(0).Rows(0).Item(0).ToString, True)), (Len(FormatFloatingAmount(ds.Tables(0).Rows(0).Item(0).ToString, True)) - 2)))

            pdfFormFields.SetField(pdfFieldPath & "D2a_2[0]", Left((FormatFloatingAmount(ds.Tables(0).Rows(0).Item(1).ToString, True)), (Len(FormatFloatingAmount(ds.Tables(0).Rows(0).Item(1).ToString, True)) - 2)))
            pdfFormFields.SetField(pdfFieldPath & "D2a_3[0]", Right((FormatFloatingAmount(FormatNumber(ds.Tables(0).Rows(0).Item(1), 2).ToString, True)), 2))

            For i As Integer = 2 To 10 Step 2
                If ds.Tables(0).Rows(0).Item(i).ToString <> 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "B20b_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(i).ToString, 0)))
                    Dim percentage As Integer
                    If i = 2 Then
                        percentage = 5
                    ElseIf i = 4 Then
                        percentage = 8
                    ElseIf i = 6 Then
                        percentage = 10
                    ElseIf i = 8 Then
                        percentage = 12
                    ElseIf i = 10 Then
                        percentage = 15
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "B20b_2[0]", CStr(percentage))
                    pdfFormFields.SetField(pdfFieldPath & "B20b_3[0]", Left((FormatFloatingAmount(FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(i + 1)), 2).ToString, True)), (Len(FormatFloatingAmount(FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(i + 1)), 2).ToString, True)) - 2)))
                    pdfFormFields.SetField(pdfFieldPath & "B20b_4[0]", Right((FormatFloatingAmount(FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(i + 1)), 2).ToString, True)), 2))
                    Exit For
                    'pdfFormFields.SetField(pdfFieldPath & "B20b_3[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(3).ToString, 2).ToString.Replace(".", "")))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "B20b_1[0]", "0")
                    pdfFormFields.SetField(pdfFieldPath & "B20b_2[0]", "0")
                    pdfFormFields.SetField(pdfFieldPath & "B20b_3[0]", "0")
                    pdfFormFields.SetField(pdfFieldPath & "B20b_4[0]", "00")
                End If
            Next

            If ds.Tables(0).Rows(0).Item(12).ToString <> 0 Then
                pdfFormFields.SetField(pdfFieldPath & "B20c_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(12).ToString, 0)))
                pdfFormFields.SetField(pdfFieldPath & "B20c_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(13).ToString, 0)))
                pdfFormFields.SetField(pdfFieldPath & "B20c_3[0]", Left((FormatFloatingAmount(FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(14)), 2).ToString, True)), (Len(FormatFloatingAmount(FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(14)), 2).ToString, True)) - 2)))
                pdfFormFields.SetField(pdfFieldPath & "B20c_4[0]", Right((FormatFloatingAmount(FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(14)), 2).ToString, True)), 2))
            Else
                pdfFormFields.SetField(pdfFieldPath & "B20c_1[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "B20c_2[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "B20c_3[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "B20c_4[0]", "00")
            End If


            'pdfFormFields.SetField(pdfFieldPath & "D2b_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(2).ToString, 0)))
            'pdfFormFields.SetField(pdfFieldPath & "D2b_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(3).ToString, 2).ToString.Replace(".", "")))
            'pdfFormFields.SetField(pdfFieldPath & "D2c_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(4).ToString, 0)))
            'pdfFormFields.SetField(pdfFieldPath & "D2c_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(5).ToString, 2).ToString.Replace(".", "")))
            'pdfFormFields.SetField(pdfFieldPath & "D2d_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(6).ToString, 0)))
            'pdfFormFields.SetField(pdfFieldPath & "D2d_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(7).ToString, 2).ToString.Replace(".", "")))
            'pdfFormFields.SetField(pdfFieldPath & "D2e_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(8).ToString, 0)))
            'pdfFormFields.SetField(pdfFieldPath & "D2e_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(9).ToString, 2).ToString.Replace(".", "")))
            'pdfFormFields.SetField(pdfFieldPath & "D2f_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(10).ToString, 0)))
            'pdfFormFields.SetField(pdfFieldPath & "D2f_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(11).ToString, 2).ToString.Replace(".", "")))
            'pdfFormFields.SetField(pdfFieldPath & "D2g_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(12).ToString, 0)))
            'pdfFormFields.SetField(pdfFieldPath & "D2g_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(13).ToString, 1).ToString.Replace(".", "")))
            'pdfFormFields.SetField(pdfFieldPath & "D2g_3[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(14).ToString, 2).ToString.Replace(".", "")))

            ds = datHandler.GetData("SELECT TCR_AMOUNT" _
                & " FROM TAX_REBATE WHERE TC_KEY=" + ComputationKey + " AND TCR_KEY=5")
            pdfFormFields.SetField(pdfFieldPath & "D4[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(0).ToString, 2).ToString.Replace(".", "")))



            ' === Part E === '
            dr = datHandler.GetDataReader("Select TC_TAX_PAYABLE, (cdbl(TC_INSTALLMENT_PAYMENT_SELF) + cdbl(TC_INSTALLMENT_PAYMENT_HW))," _
                                     & " TC_BALANCE_TAX_PAYABLE, TC_BALANCE_TAX_OVERPAID" _
                                     & " From TAX_COMPUTATION Where" _
                                     & " TC_REF_NO= '" & pdfForm.GetRefNo & "' and TC_YA= '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                If Not IsDBNull(dr("TC_TAX_PAYABLE")) Then
                    If Not String.IsNullOrEmpty(dr("TC_TAX_PAYABLE").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "E1[0]", FormatFloatingAmount(FormatNumber(CDbl(dr("TC_TAX_PAYABLE")), 2).ToString, True))
                    End If
                End If
                If Not IsDBNull(dr.Item(1)) Then
                    If Not String.IsNullOrEmpty(dr.Item(1).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "E2_1[0]", Left((FormatFloatingAmount(FormatNumber(CDbl(dr.Item(1)), 2).ToString, True)), (Len(FormatFloatingAmount(FormatNumber(CDbl(dr.Item(1)), 2).ToString, True)) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "E2_2[0]", Right((FormatFloatingAmount(FormatNumber(CDbl(dr.Item(1)), 2).ToString, True)), 2))
                    End If
                End If
                If Not IsDBNull(dr("TC_BALANCE_TAX_PAYABLE")) Then
                    If Not dr("TC_BALANCE_TAX_PAYABLE").ToString = "0.00" Then
                        pdfFormFields.SetField(pdfFieldPath & "B26_1[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "B26_2[0]", Left((FormatFloatingAmount(FormatNumber(CDbl(dr("TC_BALANCE_TAX_PAYABLE")), 2).ToString, True)), (Len(FormatFloatingAmount(FormatNumber(CDbl(dr("TC_BALANCE_TAX_PAYABLE")), 2).ToString, True)) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B26_3[0]", Right((FormatFloatingAmount(FormatNumber(CDbl(dr("TC_BALANCE_TAX_PAYABLE")), 2).ToString, True)), 2))
                    End If
                End If
                If Not IsDBNull(dr("TC_BALANCE_TAX_OVERPAID")) Then
                    If Not dr("TC_BALANCE_TAX_OVERPAID").ToString = "0.00" Then
                        pdfFormFields.SetField(pdfFieldPath & "B26_1[0]", "X")
                        pdfFormFields.SetField(pdfFieldPath & "B26_2[0]", Left((FormatFloatingAmount(FormatNumber(CDbl(dr("TC_BALANCE_TAX_OVERPAID")), 2).ToString, True)), (Len(FormatFloatingAmount(FormatNumber(CDbl(dr("TC_BALANCE_TAX_OVERPAID")), 2).ToString, True)) - 2)))
                        pdfFormFields.SetField(pdfFieldPath & "B26_3[0]", Right((FormatFloatingAmount(FormatNumber(CDbl(dr("TC_BALANCE_TAX_OVERPAID")), 2).ToString, True)), 2))
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try

    End Sub
    Private Sub Page2()

        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim prmOledb(0) As SqlParameter
        Dim strArray(1) As String

        Dim dr As SqlDataReader = Nothing, dr1 As SqlDataReader = Nothing
        Dim cSQL As String
        Dim PLKey As Long
        Dim nTotal As Double, nTotal2 As Double, nTotal3 As Double

        Dim i As Integer

        pdfFieldPath = pdfSubFormName & "Page3[0]."

        ReDim strArray(1)
        strArray = SplitText(RefName, 48)
        pdfFormFields.SetField(pdfFieldPath & "Name2_1[0]", strArray(0))
        pdfFormFields.SetField(pdfFieldPath & "Name2_2[0]", strArray(1))
        pdfFormFields.SetField(pdfFieldPath & "Ref2[0]", pdfForm.GetRefNo)

        Try

            'Part F
            dr = datHandler.GetDataReader("SELECT PY_KEY " _
                            & "FROM PRECEDING_YEAR WHERE PY_REF_NO ='" + pdfForm.GetRefNo + "' AND PY_YA ='" + pdfForm.GetYA + "'")
            If dr.Read Then
                i = 0
                dr1 = datHandler.GetDataReader("SELECT TOP 3 *" _
                                & " FROM PRECEDING_YEAR_DETAIL WHERE PY_KEY=" + dr("PY_KEY").ToString)
                While dr1.Read
                    i = i + 1
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_1[0]", dr1("PY_INCOME_TYPE").ToString.ToUpper)
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_2[0]", dr1("PY_PAYMENT_YEAR").ToString)
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_3[0]", CDbl(FormatNumber(dr1("PY_AMOUNT"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_4[0]", CDbl(FormatNumber(dr1("PY_EPF"), 0)))
                End While
                While i < 3
                    i = i + 1
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_3[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_4[0]", 0)
                End While
                dr1.Close()
            Else
                pdfFormFields.SetField(pdfFieldPath & "F1_3[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "F1_4[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "F2_3[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "F2_4[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "F3_3[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "F3_4[0]", 0)
            End If
            dr.Close()

            'Part L Profit and Loss Account
            dr = datHandler.GetDataReader("SELECT * FROM PROFIT_LOSS_ACCOUNT WHERE PL_REF_NO = '" & pdfForm.GetRefNo & "' AND PL_YA = '" & pdfForm.GetYA & "' and PL_MAINCOMPANY = '1' order by PL_KEY")
            If dr.Read Then
                cSQL = "SELECT PL_MAIN_BUSINESS, PL_KEY, PL_COMPANY" _
                & " FROM PROFIT_LOSS_ACCOUNT WHERE PL_REF_NO = '" & pdfForm.GetRefNo & "' AND PL_YA = '" & pdfForm.GetYA & "' and PL_MAINCOMPANY = '1' ORDER BY PL_KEY "
            Else
                cSQL = "SELECT PL_MAIN_BUSINESS, PL_KEY, PL_COMPANY" _
                & " FROM PROFIT_LOSS_ACCOUNT WHERE PL_REF_NO = '" & pdfForm.GetRefNo & "' AND PL_YA = '" & pdfForm.GetYA & "' ORDER BY PL_KEY "
            End If
            dr.Close()
            PnLMainBusiness = ""
            dr = datHandler.GetDataReader(cSQL)
            If dr.Read Then
                PnLMainBusiness = dr("PL_MAIN_BUSINESS")
                dr1 = datHandler.GetDataReader("SELECT BC_BUS_ENTITY, BC_CODE,BC_TYPE FROM BUSINESS_SOURCE WHERE BC_KEY = '" & pdfForm.GetRefNo & "' AND BC_YA = '" & pdfForm.GetYA & "' AND BC_BUSINESSSOURCE = '" & Trim(dr("PL_MAIN_BUSINESS")) & "'")
                If dr1.Read Then
                    pdfFormFields.SetField(pdfFieldPath & "L1[0]", dr1("BC_BUS_ENTITY").ToString.ToUpper)
                    pdfFormFields.SetField(pdfFieldPath & "L1A[0]", dr1("BC_CODE"))
                    'azham 16-feb-2016 =============================================================
                    pdfFormFields.SetField(pdfFieldPath & "ZK3[0]", dr1("BC_TYPE"))
                    'azham 16-feb-2016 =============================================================
                End If
                dr1.Close()
                PLKey = CLng(dr("PL_KEY"))
                nTotal2 = 0
                nTotal3 = 0
                'Sales
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(PL_AMOUNT as money)) FROM PL_SALES WHERE PL_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L2[0]", CDbl(FormatNumber(dr1(0), 0)).ToString)
                        nTotal3 = CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L2[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L2[0]", "0")
                End If
                dr1.Close()
                'Opening Stock
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(PL_AMOUNT as money)) FROM PL_OPENSTOCK WHERE PL_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L3[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal2 = CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L3[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L3[0]", "0")
                End If
                dr1.Close()
                'Purchase and Cost of Production
                nTotal = 0
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(PL_AMOUNT as money)) FROM PL_PURCHASE WHERE PL_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_PRODUCTION_COST WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                nTotal2 = nTotal2 + nTotal
                pdfFormFields.SetField(pdfFieldPath & "L4[0]", nTotal)
                'Closing Stock
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(PL_AMOUNT as money)) FROM PL_CLOSESTOCK WHERE PL_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L5[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal2 = nTotal2 - CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L5[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L5[0]", "0")
                End If
                dr1.Close()
                'Cost of Sales
                pdfFormFields.SetField(pdfFieldPath & "L6[0]", nTotal2)
                'Gross Profit / Loss
                nTotal3 = nTotal3 - nTotal2
                If nTotal3 < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L7_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L7_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L7_2[0]", Abs(nTotal3))
                nTotal = 0
                nTotal2 = 0
                'Other Business Income
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_OTHERBUSINESS WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = CDbl(FormatNumber(dr1(0), 0))
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L8[0]", "0")
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT [PL_KEY] FROM [PROFIT_LOSS_ACCOUNT]" _
                        & " WHERE [PL_REF_NO] = '" & pdfForm.GetRefNo & "' AND [PL_YA] = '" & pdfForm.GetYA & "' and [PL_KEY] <> " & CStr(PLKey))
                While dr1.Read()
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = nTotal2 + OtherSource_GrossProfitLoss(CLng(dr1(0)), dr("PL_COMPANY").ToString, PnLMainBusiness, pdfForm.GetRefNo, pdfForm.GetYA)
                    End If
                End While
                dr1.Close()
                pdfFormFields.SetField(pdfFieldPath & "L8[0]", nTotal2)
                nTotal = nTotal2

                'Dividends
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_NONBUSINESS WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 47")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L9[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L9[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L9[0]", "0")
                End If
                dr1.Close()
                'Interest and discounts
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_NONBUSINESS WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 50")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L10[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L10[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L10[0]", "0")
                End If
                dr1.Close()
                'Rents, royalties and premiums
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_NONBUSINESS WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE BETWEEN 48 and 49")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L11[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L11[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L11[0]", "0")
                End If
                dr1.Close()
                nTotal2 = 0
                'Other Income
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_NONBUSINESS WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 51")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_NONTAXABLE WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = nTotal2 + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                pdfFormFields.SetField(pdfFieldPath & "L12[0]", nTotal2)
                nTotal = nTotal + nTotal2
                pdfFormFields.SetField(pdfFieldPath & "L13[0]", nTotal)
                nTotal3 = nTotal3 + nTotal
                nTotal = 0
                'Loan interest
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 11")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L14[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L14[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L14[0]", "0")
                End If
                dr1.Close()
                'Salaries and wages
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 12")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L15[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L15[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L15[0]", "0")
                End If
                dr1.Close()
                'Rental / Lease
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 13")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L16[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L16[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L16[0]", "0")
                End If
                dr1.Close()
                'Contracts and subcontracts
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 14")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L17[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L17[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L17[0]", "0")
                End If
                dr1.Close()
                'Commissions
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 15")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L18[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L18[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L18[0]", "0")
                End If
                dr1.Close()
                'Bad debts
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 16")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L19[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L19[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L19[0]", "0")
                End If
                dr1.Close()
                'Travelling and transport
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 17")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L20[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L20[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L20[0]", "0")
                End If
                dr1.Close()
                'Repair and maintenance
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 52")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L21[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L21[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L21[0]", "0")
                End If
                dr1.Close()
                'Promotion and advertisement
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 53")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L22[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L22[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L22[0]", "0")
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(EXA_AMOUNT) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 54")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "D7A_10[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "D7A_10[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "D7A_10[0]", "0")
                End If
                dr1.Close()

                nTotal2 = 0
                'Other expenses
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND (EXA_PLTYPE between 18 and 20 or EXA_PLTYPE = 46)")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXP_NONALLOWLOSS WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = nTotal2 + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXP_NONALLOWEXPEND WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = nTotal2 + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXP_PERSONAL WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = nTotal2 + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                pdfFormFields.SetField(pdfFieldPath & "L23[0]", nTotal2)
                nTotal = nTotal + nTotal2
                pdfFormFields.SetField(pdfFieldPath & "L24[0]", nTotal)
                nTotal3 = nTotal3 - nTotal
                If nTotal3 < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L25_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L25_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L25_2[0]", Abs(nTotal3))
                nTotal = 0
                'Non-allowable expenses
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & " AND [EXA_DEDUCTIBLE]='No'")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXP_NONALLOWEXPEND WHERE EXA_KEY = " & CStr(PLKey) & " AND [EXA_DEDUCTIBLE]='No'")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXP_PERSONAL WHERE EXA_KEY = " & CStr(PLKey) & " AND [EXA_DEDUCTIBLE]='No'")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_PRODUCTION_COST WHERE EXA_KEY = " & CStr(PLKey) & "AND [EXA_DEDUCTIBLE]='No' and (EXA_PLTYPE = 43 or EXA_PLTYPE = 45)")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                pdfFormFields.SetField(pdfFieldPath & "L26[0]", nTotal)
            Else
                pdfFormFields.SetField(pdfFieldPath & "L1A[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L2[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L3[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L4[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L5[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L6[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L7_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L7_2[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L8[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L9[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L10[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L11[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L12[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L13[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L14[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L15[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L16[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L17[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L18[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L19[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L20[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L21[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L22[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L23[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L24[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L25_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L25_2[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L26[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "D7A_10[0]", "0")
            End If

            'Part L Balance Sheet
            dr = datHandler.GetDataReader("SELECT *" _
                & " FROM BALANCE_SHEET WHERE BS_REF_NO = '" & pdfForm.GetRefNo & "' AND BS_YA = '" & pdfForm.GetYA & "' AND [BS_SOURCENO] = '" & Trim(PnLMainBusiness) + "' ORDER BY BS_SOURCENO")
            If dr.Read Then
                pdfFormFields.SetField(pdfFieldPath & "L27[0]", CDbl(FormatNumber(dr("BS_LAND"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L28[0]", CDbl(FormatNumber(dr("BS_MACHINERY"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L29[0]", CDbl(FormatNumber(dr("BS_TRANSPORT"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L30[0]", CDbl(FormatNumber(dr("BS_OTH_FA"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L31[0]", CDbl(FormatNumber(dr("BS_TOT_FA"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L32[0]", CDbl(FormatNumber(dr("BS_INVESTMENT"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L33[0]", CDbl(FormatNumber(dr("BS_STOCK"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L34[0]", CDbl(FormatNumber(dr("BS_TRADE_DEBTORS"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L35[0]", CDbl(FormatNumber(dr("BS_OTH_DEBTORS"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L36[0]", CDbl(FormatNumber(dr("BS_CASH"), 0)))
                If CDbl(FormatNumber(dr("BS_BANK"), 0)) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L37_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L37_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L37_2[0]", Abs(CDbl(FormatNumber(dr("BS_BANK"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "L38[0]", CDbl(FormatNumber(dr("BS_OTH_CA"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L39[0]", CDbl(FormatNumber(dr("BS_TOT_CA"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L40[0]", CDbl(FormatNumber(dr("BS_TOT_ASSETS"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L41[0]", CDbl(FormatNumber(dr("BS_LOAN"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L42[0]", CDbl(FormatNumber(dr("BS_TRADE_CR"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L43[0]", (CDbl(FormatNumber(dr("BS_OTHER_CR"), 0)) + CDbl(FormatNumber(dr("BS_OTH_LIAB"), 0)) + CDbl(FormatNumber(dr("BS_LT_LIAB"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "L44[0]", CDbl(FormatNumber(dr("BS_TOT_LIAB"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L45[0]", CDbl(FormatNumber(dr("BS_CAPITALACCOUNT"), 0)))
                If CDbl(FormatNumber(dr("BS_BROUGHT_FORWARD"), 0)) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L46_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L46_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L46_2[0]", Abs(CDbl(FormatNumber(dr("BS_BROUGHT_FORWARD"), 0))))
                If CDbl(FormatNumber(dr("BS_CY_PROFITLOSS"), 0)) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L47_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L47_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L47_2[0]", Abs(CDbl(FormatNumber(dr("BS_CY_PROFITLOSS"), 0))))
                If (CDbl(FormatNumber(dr("BS_CAP_CONTRIBUTION"), 0)) - CDbl(FormatNumber(dr("BS_DRAWING"), 0))) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L48_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L48_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L48_2[0]", Abs(CDbl(FormatNumber(dr("BS_CAP_CONTRIBUTION"), 0)) - CDbl(FormatNumber(dr("BS_DRAWING"), 0))))
                If CDbl(FormatNumber(dr("BS_CARRIED_FORWARD"), 0)) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L49_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L49_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L49_2[0]", Abs(CDbl(FormatNumber(dr("BS_CARRIED_FORWARD"), 0))))
            Else
                pdfFormFields.SetField(pdfFieldPath & "L27[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L28[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L29[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L30[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L31[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L32[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L33[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L34[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L35[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L36[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L37_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L37_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L38[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L39[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L40[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L41[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L42[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L43[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L44[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L45[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L46_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L46_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L47_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L47_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L48_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L48_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L49_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L49_2[0]", 0)
            End If

            'Part A
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            ds = datHandler.GetData("SELECT (TP_REG_ADD_LINE1 + ', ' + TP_REG_ADD_LINE2 + ', ' + TP_REG_ADD_LINE3)," _
                            & " TP_REG_POSTCODE, TP_REG_CITY, TP_REG_STATE," _
                            & " (TP_COM_ADD_LINE1 + ', ' + TP_COM_ADD_LINE2 + ', ' + TP_COM_ADD_LINE3), TP_COM_POSTCODE, TP_COM_CITY, TP_COM_STATE," _
                            & " TP_TEL1, TP_TEL2, TP_EMAIL, TP_BANK, TP_BANK_ACC, TP_EMPLOYERNAME, (TP_EMPLOYER_NO2 + TP_EMPLOYER_NO3)," _
                            & " TP_HW_NAME, TP_HW_REF_NO_PREFIX, TP_HW_REF_NO1," _
                            & " (TP_HW_IC_NEW1 + TP_HW_IC_NEW2 + TP_HW_IC_NEW3), TP_HW_IC_OLD, TP_HW_POLICE_NO," _
                            & " TP_HW_ARMY_NO, TP_HW_PASSPORT_NO, TP_PASSWPORTDUEDATE2" _
                            & " FROM TAXP_PROFILE WHERE (TP_REF_NO1 + TP_REF_NO2 + TP_REF_NO3)=@ref_no", prmOledb)

            ReDim strArray(2)
            If ds.Tables(0).Rows(0).Item(0).ToString() <> ",, " Then
                strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().Replace(",,", ",").Replace(", ,", ",").ToUpper, 68)
                pdfFormFields.SetField(pdfFieldPath & "A10_1[0]", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "A10_2[0]", strArray(1))
                'pdfFormFields.SetField(pdfFieldPath & "A10_3[0]", strArray(2))
                pdfFormFields.SetField(pdfFieldPath & "A10_4[0]", ds.Tables(0).Rows(0).Item(1).ToString)
                pdfFormFields.SetField(pdfFieldPath & "A10_5[0]", ds.Tables(0).Rows(0).Item(2).ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "A10_6[0]", ds.Tables(0).Rows(0).Item(3).ToString.ToUpper)
            Else
                pdfFormFields.SetField(pdfFieldPath & "A10_1[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A10_2[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A10_3[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A10_4[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A10_5[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A10_6[0]", "---")
            End If
            ReDim strArray(2)
            If ds.Tables(0).Rows(0).Item(4).ToString() <> ",, " Then
                strArray = SplitText(ds.Tables(0).Rows(0).Item(4).ToString().Replace(",,", ",").Replace(", ,", ",").ToUpper, 26)
                pdfFormFields.SetField(pdfFieldPath & "A11_1[0]", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "A11_2[0]", strArray(1))
                pdfFormFields.SetField(pdfFieldPath & "A11_3[0]", strArray(2))
                pdfFormFields.SetField(pdfFieldPath & "A11_4[0]", ds.Tables(0).Rows(0).Item(5).ToString)
                pdfFormFields.SetField(pdfFieldPath & "A11_5[0]", ds.Tables(0).Rows(0).Item(6).ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "A11_6[0]", ds.Tables(0).Rows(0).Item(7).ToString.ToUpper)
            Else
                pdfFormFields.SetField(pdfFieldPath & "A11_1[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A11_2[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A11_3[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A11_4[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A11_5[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A11_6[0]", "---")
            End If
            If ds.Tables(0).Rows(0).Item(8).ToString() <> "" And ds.Tables(0).Rows(0).Item(9).ToString() <> "" Then
                pdfFormFields.SetField(pdfFieldPath & "A12[0]", FormatPhoneNumber(ds.Tables(0).Rows(0).Item(8).ToString, ds.Tables(0).Rows(0).Item(9).ToString, " ", " "))
            Else
                pdfFormFields.SetField(pdfFieldPath & "A12[0]", "---")
            End If
            If ds.Tables(0).Rows(0).Item(10).ToString() <> "" Then
                pdfFormFields.SetField(pdfFieldPath & "A13[0]", ds.Tables(0).Rows(0).Item(10).ToString)
            Else
                pdfFormFields.SetField(pdfFieldPath & "A13[0]", "---")
            End If
            'If ds.Tables(0).Rows(0).Item(11).ToString() <> "" Then
            '    pdfFormFields.SetField(pdfFieldPath & "A14[0]", ds.Tables(0).Rows(0).Item(11).ToString.ToUpper)
            'Else
            '    pdfFormFields.SetField(pdfFieldPath & "A14[0]", "---")
            'End If
            If ds.Tables(0).Rows(0).Item(11).ToString() <> "" Then
                strArray = SplitText(ds.Tables(0).Rows(0).Item(11).ToString().ToUpper, 26)
                pdfFormFields.SetField(pdfFieldPath & "A15[0]", strArray(0))
            Else
                pdfFormFields.SetField(pdfFieldPath & "A15[0]", "---")
            End If
            If ds.Tables(0).Rows(0).Item(12).ToString() <> "" Then
                pdfFormFields.SetField(pdfFieldPath & "A16[0]", ds.Tables(0).Rows(0).Item(12).ToString)
            Else
                pdfFormFields.SetField(pdfFieldPath & "A16[0]", "---")
            End If
            ReDim strArray(1)
            If ds.Tables(0).Rows(0).Item(13).ToString() <> "" Then
                strArray = SplitText(ds.Tables(0).Rows(0).Item(13).ToString().ToUpper, 26)
                pdfFormFields.SetField(pdfFieldPath & "A17_1[0]", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "A17_2[0]", strArray(1))
            Else
                pdfFormFields.SetField(pdfFieldPath & "A17_1[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "A17_2[0]", "---")
            End If
            If ds.Tables(0).Rows(0).Item(14).ToString() <> "" Then
                pdfFormFields.SetField(pdfFieldPath & "A18[0]", ds.Tables(0).Rows(0).Item(14).ToString)
            Else
                pdfFormFields.SetField(pdfFieldPath & "A18[0]", "---")
            End If

            'Part B
            'ReDim strArray(1)
            'strArray = SplitText(ds.Tables(0).Rows(0).Item(15).ToString().ToUpper, 28)
            pdfFormFields.SetField(pdfFieldPath & "C1[0]", ds.Tables(0).Rows(0).Item(15).ToString)
            'pdfFormFields.SetField(pdfFieldPath & "B1_2[0]", strArray(1))
            'pdfFormFields.SetField(pdfFieldPath & "B2_1[0]", ds.Tables(0).Rows(0).Item(16).ToString)
            'pdfFormFields.SetField(pdfFieldPath & "B2_2[0]", ds.Tables(0).Rows(0).Item(17).ToString)
            pdfFormFields.SetField(pdfFieldPath & "C2[0]", FormatICNumber(ds.Tables(0).Rows(0).Item(18).ToString))
            'pdfFormFields.SetField(pdfFieldPath & "B4[0]", ds.Tables(0).Rows(0).Item(19).ToString)
            'pdfFormFields.SetField(pdfFieldPath & "B5[0]", ds.Tables(0).Rows(0).Item(20).ToString)
            'pdfFormFields.SetField(pdfFieldPath & "B6[0]", ds.Tables(0).Rows(0).Item(21).ToString)
            'pdfFormFields.SetField(pdfFieldPath & "B7[0]", ds.Tables(0).Rows(0).Item(22).ToString)
            If Not IsDBNull(ds.Tables(0).Rows(0).Item(23)) Then
                pdfFormFields.SetField(pdfFieldPath & "C4[0]", FormatDateSlash(ds.Tables(0).Rows(0).Item(23)))
            End If

            'Master Data
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            ds = datHandler.GetData("SELECT TP_HW_LAST_PASSPORT_NO, TP_HW_DOB, TP_BWA FROM TAXP_PROFILE2 WHERE" _
                        & " TP_REF_NO= ?", prmOledb)
            pdfFormFields.SetField(pdfFieldPath & "A14[0]", ds.Tables(0).Rows(0).Item(2).ToString)
            pdfFormFields.SetField(pdfFieldPath & "C5[0]", ds.Tables(0).Rows(0).Item(0).ToString)

            If Not IsDBNull(ds.Tables(0).Rows(0).Item(1)) Then
                pdfFormFields.SetField(pdfFieldPath & "C6[0]", FormatDateSlash(ds.Tables(0).Rows(0).Item(1)))
            Else
                pdfFormFields.SetField(pdfFieldPath & "C6[0]", "")
            End If



        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try
    End Sub
    Private Sub Page3()
        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr2 As SqlDataReader = Nothing
        Dim prmOledb(0) As SqlParameter
        Dim strArray(1) As String
        Dim intCounter As Integer = 1
        Dim intNumberRecord As Integer = 0
        Dim dblTotalIncome As Double = 0
        Dim dblRentalIncome As Double = 0

        pdfFieldPath = pdfSubFormName & "Page4[0]."
        ReDim strArray(1)
        strArray = SplitText(RefName, 48)
        pdfFormFields.SetField(pdfFieldPath & "Name3_1[0]", strArray(0))
        pdfFormFields.SetField(pdfFieldPath & "Name3_2[0]", strArray(1))
        pdfFormFields.SetField(pdfFieldPath & "Ref3[0]", pdfForm.GetRefNo)

        Try

            ReDim strArray(1)
            dr = datHandler.GetDataReader("SELECT * FROM [TAXA_PROFILE] Where [TA_KEY] =" & pdfForm.GetTaxAgent)
            If dr.Read() Then
                If Not IsDBNull(dr("TA_CO_NAME")) Then
                    If Not String.IsNullOrEmpty(dr("TA_CO_NAME")) Then
                        strArray = SplitText(dr("TA_CO_NAME").ToString, 32)
                        pdfFormFields.SetField(pdfFieldPath & "NyataA_1", strArray(0))
                        pdfFormFields.SetField(pdfFieldPath & "NyataA_2", strArray(1))
                    End If
                End If
                pdfFormFields.SetField(pdfFieldPath & "Nyatab", FormatPhoneNumber("", dr("TA_TEL_NO").ToString, "", dr("TA_MOBILE").ToString))
                If Not IsDBNull(dr("TA_LICENSE")) Then
                    pdfFormFields.SetField(pdfFieldPath & "Nyatac", dr("TA_LICENSE"))
                End If
                pdfFormFields.SetField(pdfFieldPath & "NyataTarikh", FormatDate(Now))
            End If
            dr.Close()

            'Part C
            'Count number of adjuster income record"
            dr = datHandler.GetDataReader("select count(adj_key) from income_adjusted where" _
                        & " adj_ref_no= '" & pdfForm.GetRefNo & "' and adj_ya='" & pdfForm.GetYA & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then
                    intNumberRecord = CInt(dr.Item(0))
                End If
            End If
            dr.Close()

            'Get rental income"
            dr = datHandler.GetDataReader("select os_rt_sec4A_rental from income_othersource where" _
                        & " os_ref_no='" & pdfForm.GetRefNo & "' and os_ya='" & pdfForm.GetYA & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then
                    dblRentalIncome = CDbl(dr.Item(0))
                End If
            End If
            dr.Close()


            'C1 , C2 
            dr = datHandler.GetDataReader("select top 2 adjsi_net_stat_income , adj_business from income_adjusted where" _
                                          & " adj_ref_no='" & pdfForm.GetRefNo & "' and adj_ya= '" & pdfForm.GetYA & "' order by adj_business")
            pdfFormFields.SetField(pdfFieldPath & "C1_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "C2_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "C3_2[0]", 0)
            'If dr.Read Then
            While dr.Read()

                dr2 = datHandler.GetDataReader("select bc_code from business_source where " _
                            & " bc_key='" & pdfForm.GetRefNo & "' and bc_ya='" & pdfForm.GetYA & "'" _
                            & " and bc_businesssource='" & dr("adj_business") & "'")

                If dr2.Read() Then
                    If Not IsDBNull(dr2("bc_code")) Then
                        pdfFormFields.SetField(pdfFieldPath & "C" & intCounter.ToString & "_1[0]", dr2("bc_code").ToString)
                    End If
                    If Not IsDBNull(dr("adjsi_net_stat_income")) Then
                        pdfFormFields.SetField(pdfFieldPath & "C" & intCounter.ToString & "_2[0]", CDbl(FormatNumber((dr("adjsi_net_stat_income").ToString), 0)))
                        dblTotalIncome = dblTotalIncome + CDbl(dr("adjsi_net_stat_income"))
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "C1_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "C2_2[0]", 0)
                End If
                intCounter = intCounter + 1
                dr2.Close()
            End While
            'Else
            'End If
            dr.Close()

            'if rental income
            'C3
            If intNumberRecord <= 2 Then
                If dblRentalIncome > 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "C" & (intNumberRecord + 1).ToString & "_1[0]", "70102")
                    pdfFormFields.SetField(pdfFieldPath & "C" & (intNumberRecord + 1).ToString & "_2[0]", CDbl(FormatNumber((dblRentalIncome.ToString), 0)))
                End If
            Else
                If intNumberRecord = 3 And dblRentalIncome = 0 Then
                    dr = datHandler.GetDataReader("select top 3 adjsi_net_stat_income , adj_business from income_adjusted where" _
                                & " adj_ref_no='" & pdfForm.GetRefNo & "' and adj_ya= '" & pdfForm.GetYA & "' order by adj_business desc")
                    If dr.Read() Then
                        dr2 = datHandler.GetDataReader("select bc_code from business_source where " _
                                    & " bc_key='" & pdfForm.GetRefNo & "' and bc_ya='" & pdfForm.GetYA & "'" _
                                    & " and BC_BUSINESSSOURCE='" & dr("adj_business").ToString & "'")
                        If dr2.Read() Then
                            If Not IsDBNull(dr2("bc_code")) Then
                                pdfFormFields.SetField(pdfFieldPath & "C3_1[0]", dr2("bc_code").ToString)
                            End If
                            If Not IsDBNull(dr("adjsi_net_stat_income")) Then
                                pdfFormFields.SetField(pdfFieldPath & "C3_2[0]", CDbl(FormatNumber((dr("adjsi_net_stat_income").ToString), 0)))
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "C2_2[0]", 0)
                            End If
                        End If
                        dr2.Close()
                    End If
                    dr.Close()

                Else
                    dr = datHandler.GetDataReader("select SUM(cast(adjsi_net_stat_income as money)) from income_adjusted where " _
                                   & "adj_ref_no ='" & pdfForm.GetRefNo & "' and adj_ya= '" & pdfForm.GetYA & "'")
                    If dr.Read() Then
                        If Not IsDBNull(dr.Item(0)) Then
                            dblTotalIncome = CDbl(dr.Item(0)) - dblTotalIncome
                        End If
                    End If
                    dblTotalIncome = dblTotalIncome + dblRentalIncome
                    If dblTotalIncome > 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "C3_2[0]", CDbl(FormatNumber((dblTotalIncome), 0)))
                    End If
                    dr.Close()
                End If
            End If

            intNumberRecord = 0
            intCounter = 4
            dblTotalIncome = 0

            'C4 , C5, C6
            dr = datHandler.GetDataReader("select count(pn_key) from income_partnership where " _
                                 & " pn_ref_no='" & pdfForm.GetRefNo & "'and pn_ya= '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then
                    intNumberRecord = CInt(dr.Item(0))
                End If
            End If
            dr.Close()

            If intNumberRecord = 3 Then
                dr = datHandler.GetDataReader("select top 3 ps_sch_7a_stat_income , ps_source from income_partnership where " _
                                & " pn_ref_no='" & pdfForm.GetRefNo & "'and pn_ya= '" & pdfForm.GetYA & "'" _
                                & " order by ps_source")
            Else
                dr = datHandler.GetDataReader("select top 2 ps_sch_7a_stat_income , ps_source from income_partnership where " _
                                & " pn_ref_no='" & pdfForm.GetRefNo & "'and pn_ya= '" & pdfForm.GetYA & "'" _
                                & " order by ps_source")
            End If
            pdfFormFields.SetField(pdfFieldPath & "C4_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "C5_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "C6_2[0]", 0)
            'If dr.Read Then
            While dr.Read()
                dr2 = datHandler.GetDataReader("select (ps_file_no2 + ps_file_no3) from taxp_partnership where " _
                                & "ps_key='" & pdfForm.GetRefNo & "'and ps_ya='" & pdfForm.GetYA & "' and " _
                                & "ps_sourceno=" & dr("ps_source"))

                If dr2.Read Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "C" & intCounter.ToString & "_1[0]", CDbl(FormatNumber(dr2.Item(0), 0)))
                    End If
                End If
                If Not IsDBNull(dr("ps_sch_7a_stat_income")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C" & intCounter.ToString & "_2[0]", CDbl(FormatNumber((dr("ps_sch_7a_stat_income")), 0)))
                    dblTotalIncome = dblTotalIncome + CDbl(dr("ps_sch_7a_stat_income"))
                End If
                intCounter = intCounter + 1
                dr2.Close()
            End While
            'Else
            'End If
            dr.Close()

            If intNumberRecord > 3 Then
                dr = datHandler.GetDataReader("select SUM(cast(ps_sch_7a_stat_income as money)) from income_partnership where " _
                                   & " pn_ref_no='" & pdfForm.GetRefNo & "'and pn_ya= '" & pdfForm.GetYA & "'")

                If dr.Read() Then

                    If IsDBNull(dr.Item(0)) Or String.IsNullOrEmpty(dr.Item(0)) Then
                        dblTotalIncome = 0 - dblTotalIncome
                    Else
                        dblTotalIncome = dr.Item(0) - dblTotalIncome
                    End If
                    If dblTotalIncome >= 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "C6_2[0]", CDbl(FormatNumber(dblTotalIncome, 0)))
                    End If
                End If
                dr.Close()
            End If


            'C7 - C20
            dr = datHandler.GetDataReader("Select TC_STATUTORY_INCOME, TC_BUSINESSLOSS_BF, TC_AGGREGATE_BUS_INCOME," _
                                & " TC_EMPLOYMENT_INCOME, TC_DIVIDEND, (cdbl(TC_INTEREST) + cdbl(TC_DISCOUNT)), " _
                                & " (cdbl(TC_RENTAL_ROYALTY)+cdbl(TC_PREMIUM)), TC_PENSION_AND_ETC," _
                                & " (cdbl(TC_OTHER_GAIN_PROFIT) + cdbl(TC_SEC4A)), TC_ADDITION_43," _
                                & " TC_AGGREGATE_OTHER_SRC, TC_AGGREGATE_INCOME, TC_BUSINESSLOSS_CY," _
                                & " TC_TOTAL1, TC_EXEMPT_CLAIM, TC_EXEMPT_COUNTRY, TC_EXEMPT_AMOUNT from tax_computation where" _
                                & " tc_ref_no ='" & pdfForm.GetRefNo & "' and tc_ya ='" & pdfForm.GetYA & "'")
            'weihong TC_EXEMPT_AMOUNT
            If dr.Read() Then
                If Not IsDBNull(dr("TC_STATUTORY_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C7[0]", CDbl(FormatNumber((dr("TC_STATUTORY_INCOME")), 0)))
                End If
                If Not IsDBNull(dr("TC_BUSINESSLOSS_BF")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C8[0]", CDbl(FormatNumber((dr("TC_BUSINESSLOSS_BF")), 0)))
                End If
                If Not IsDBNull(dr("TC_AGGREGATE_BUS_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C9[0]", CDbl(FormatNumber((dr("TC_AGGREGATE_BUS_INCOME")), 0)))
                End If
                If CDbl(dr("TC_EXEMPT_CLAIM")) > 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "C10_1[0]", dr("TC_EXEMPT_CLAIM"))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "C10_1[0]", "")
                End If
                If Not IsDBNull(dr("TC_EXEMPT_COUNTRY")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C10_2[0]", dr("TC_EXEMPT_COUNTRY").ToString)
                End If
                If Not IsDBNull(dr("TC_EMPLOYMENT_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C10_3[0]", CDbl(FormatNumber((dr("TC_EMPLOYMENT_INCOME")), 0)))
                End If
                'weihong
                If Not IsDBNull(dr("TC_EXEMPT_AMOUNT")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C10_4[0]", CDbl(FormatNumber((dr("TC_EXEMPT_AMOUNT")), 0)))
                End If
                'weihong
                If Not IsDBNull(dr("TC_DIVIDEND")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C11[0]", CDbl(FormatNumber((dr("TC_DIVIDEND")), 0)))
                End If
                If Not IsDBNull(dr.Item(5)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C12[0]", CDbl(FormatNumber((dr.Item(5).ToString), 0)))
                End If
                If Not IsDBNull(dr.Item(6)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C13[0]", CDbl(FormatNumber((dr.Item(6).ToString), 0)))
                End If
                If Not IsDBNull(dr("TC_PENSION_AND_ETC")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C14[0]", CDbl(FormatNumber((dr("TC_PENSION_AND_ETC")), 0)))
                End If
                If Not IsDBNull(dr.Item(8)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C15[0]", CDbl(FormatNumber((dr.Item(8).ToString), 0)))
                End If
                If Not IsDBNull(dr("TC_ADDITION_43")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C16[0]", CDbl(FormatNumber((dr("TC_ADDITION_43")), 0)))
                End If
                If Not IsDBNull(dr("TC_AGGREGATE_OTHER_SRC")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C17[0]", CDbl(FormatNumber((dr("TC_AGGREGATE_OTHER_SRC")), 0)))
                End If
                If Not IsDBNull(dr("TC_AGGREGATE_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C18[0]", CDbl(FormatNumber((dr("TC_AGGREGATE_INCOME")), 0)))
                End If
                If Not IsDBNull(dr("TC_BUSINESSLOSS_CY")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C19[0]", CDbl(FormatNumber((dr("TC_BUSINESSLOSS_CY")), 0)))
                End If
                If Not IsDBNull(dr("TC_TOTAL1")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C20[0]", CDbl(FormatNumber((dr("TC_TOTAL1")), 0)))
                End If
            Else

            End If
            dr.Close()

            'azham 18-feb-2016 ======================
            'ZK========================================================================================
            strSQL = "select DP_REF_NO,DP_DISPOSAL,DP_DECLARE FROM DISPOSAL WHERE DP_REF_NO= '" & pdfForm.GetRefNo & "'"
            ds = datHandler.GetData(strSQL)
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                If IsDBNull(ds.Tables(0).Rows(0)("DP_DISPOSAL")) = False OrElse ds.Tables(0).Rows(0)("DP_DISPOSAL") <> "" Then
                    If ds.Tables(0).Rows(0)("DP_DISPOSAL") = "Yes" Then
                        pdfFormFields.SetField(pdfFieldPath & "ZK1[0]", "1")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "ZK1[0]", "2")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "ZK1[0]", "2")
                End If

                If IsDBNull(ds.Tables(0).Rows(0)("DP_DECLARE")) = False OrElse ds.Tables(0).Rows(0)("DP_DECLARE") <> "" Then
                    If ds.Tables(0).Rows(0)("DP_DECLARE") = "Yes" Then
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
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try

    End Sub
    Private Sub Page4()
        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr2 As SqlDataReader = Nothing
        Dim dblTotalValue As Double = 0
        Dim nTotal As Double = 0
        pdfFieldPath = pdfSubFormName & "Page6[0]."
        pdfFormFields.SetField(pdfFieldPath & "Nama6[0]", RefName)
        pdfFormFields.SetField(pdfFieldPath & "Ruj6[0]", pdfForm.GetRefNo)

        Try
            'Part C
            dr = datHandler.GetDataReader("select TC_KEY, TC_AGGREGATE_INCOME, TC_PROSPECTING, TC_QUALIFYING_AG_EXP, TC_TOTAL2," _
                                    & " TC_4, TC_3, TC_TOTAL_INCOME_2, TC_INCOME_TRANSFER_FROM_HW, TC_TOTAL_INCOME_3" _
                                    & " from tax_computation where" _
                                    & " tc_ref_no ='" & pdfForm.GetRefNo & "' and tc_ya ='" & pdfForm.GetYA & "'")
            If dr.Read Then

                If Not IsDBNull(dr("TC_PROSPECTING")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C21[0]", CDbl(FormatNumber((dr("TC_PROSPECTING")), 0)))
                End If
                If Not IsDBNull(dr("TC_TOTAL2")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C22[0]", CDbl(FormatNumber((dr("TC_TOTAL2")), 0)))
                End If

                'C23 - 'C30
                dr2 = datHandler.GetDataReader("select TCG_KEY, TCG_AMOUNT" _
                                        & " from tax_gifts where" _
                                        & " tc_key =" & dr("TC_KEY"))
                Do While dr2.Read()
                    If Not IsDBNull(dr2("TCG_KEY")) Then

                        Select Case dr2("TCG_KEY")
                            Case "9"
                                pdfFormFields.SetField(pdfFieldPath & "C23[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                            Case "1"
                                pdfFormFields.SetField(pdfFieldPath & "C23A[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "7"
                                pdfFormFields.SetField(pdfFieldPath & "C24[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "8"
                                pdfFormFields.SetField(pdfFieldPath & "C25_1[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "2"
                                pdfFormFields.SetField(pdfFieldPath & "C26[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                            Case "3"
                                pdfFormFields.SetField(pdfFieldPath & "C27[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                            Case "4"
                                pdfFormFields.SetField(pdfFieldPath & "C28[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                            Case "5"
                                pdfFormFields.SetField(pdfFieldPath & "C29[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))
                            Case "6"
                                pdfFormFields.SetField(pdfFieldPath & "C30[0]", CDbl(FormatNumber((dr2("TCG_AMOUNT")), 0)))

                        End Select
                    End If
                Loop
                dr2.Close()

                'C26 Total restrict to 7% of C28
                If Not IsDBNull(dr("TC_AGGREGATE_INCOME")) Then
                    If Not String.IsNullOrEmpty(dr("TC_AGGREGATE_INCOME").ToString) Then
                        If dblTotalValue >= CDbl(dr("TC_AGGREGATE_INCOME")) * 0.07 Then
                            pdfFormFields.SetField(pdfFieldPath & "C25_2[0]", CDbl(FormatNumber(((CDbl(dr("TC_AGGREGATE_INCOME")) * 0.07).ToString), 0)))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "C25_2[0]", CDbl(FormatNumber((dblTotalValue.ToString), 0)))
                        End If
                    End If
                End If

                'C31 - C35_2
                'NGOHCS B2010.2
                dr2 = datHandler.GetDataReader("select SUM(cast(TCG_AMOUNT as money))" _
                                       & " from tax_gifts where" _
                                       & " tc_key =" & dr("TC_KEY"))
                If dr2.Read() Then
                    If Not IsDBNull(dr.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "C31[0]", CDbl(dr("TC_TOTAL2")) - CDbl(dr2.Item(0)))
                        nTotal = nTotal + (CDbl(dr("TC_TOTAL2")) - CDbl(dr2.Item(0)))
                    End If
                End If
                dr2.Close()
                'NGOHCS B2010.2 END
                'If Not IsDBNull(dr("TC_4")) Then
                '    pdfFormFields.SetField(pdfFieldPath & "C31[0]", CDbl(FormatNumber((dr("TC_4")), 0)))
                '    nTotal = CDbl(dr("TC_4"))
                'End If

                If Not IsDBNull(dr("TC_3")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C32[0]", CDbl(FormatNumber((dr("TC_3")), 0)))
                    nTotal = nTotal + CDbl(dr("TC_3"))
                End If
                dr2 = datHandler.GetDataReader("select INTEREST, ROYALTIES, SECTION4A, OTHERINCOME" _
                                                        & " from CHARGEABLE_INCOME where" _
                                                        & " tc_key =" & dr("TC_KEY"))
                If dr2.Read Then
                    If Not IsDBNull(dr2("INTEREST")) Then
                        pdfFormFields.SetField(pdfFieldPath & "C33a[0]", CDbl(FormatNumber((dr2("INTEREST")), 0)))
                        nTotal = nTotal + CDbl(dr2("INTEREST"))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "C33a[0]", 0)
                    End If
                    If Not IsDBNull(dr2("ROYALTIES")) Then
                        pdfFormFields.SetField(pdfFieldPath & "C33b[0]", CDbl(FormatNumber((dr2("ROYALTIES")), 0)))
                        nTotal = nTotal + CDbl(dr2("ROYALTIES"))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "C33b[0]", 0)
                    End If
                    If Not IsDBNull(dr2("SECTION4A")) Then
                        pdfFormFields.SetField(pdfFieldPath & "C33c[0]", CDbl(FormatNumber((dr2("SECTION4A")), 0)))
                        nTotal = nTotal + CDbl(dr2("SECTION4A"))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "C33c[0]", 0)
                    End If
                    If Not IsDBNull(dr2("OTHERINCOME")) Then
                        pdfFormFields.SetField(pdfFieldPath & "C33d_1[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "C33d_2[0]", CDbl(FormatNumber((dr2("OTHERINCOME")), 0)))
                        nTotal = nTotal + CDbl(dr2("OTHERINCOME"))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "C33d_1[0]", "")
                        pdfFormFields.SetField(pdfFieldPath & "C33d_2[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "C33a[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "C33b[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "C33c[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "C33d_1[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "C33d_2[0]", 0)
                End If
                dr2.Close()
                pdfFormFields.SetField(pdfFieldPath & "C34[0]", nTotal)

                If Not IsDBNull(dr("TC_INCOME_TRANSFER_FROM_HW")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C35_2[0]", CDbl(FormatNumber((dr("TC_INCOME_TRANSFER_FROM_HW")), 0)))
                End If

                'C35_1
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
                    If Not IsDBNull(dr2("TP_HW_TYPEOFINCOME")) And Not String.IsNullOrEmpty(dr2("TP_HW_TYPEOFINCOME").ToString) Then
                        If dr2("TP_TYPE_ASSESSMENT") = "1" Then
                            If (dr2("TP_GENDER") = "1" And dr2("TP_ASSESSMENTON") = "1") Or _
                                (dr2("TP_GENDER") = "2" And dr2("TP_ASSESSMENTON") = "2") Then
                                If dr2("TP_HW_TYPEOFINCOME").ToString = "1" Or boolWithBusiness = True Then
                                    pdfFormFields.SetField(pdfFieldPath & "C35_1[0]", "1")
                                Else
                                    pdfFormFields.SetField(pdfFieldPath & "C35_1[0]", "2")
                                End If
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "C35_1[0]", "")
                            End If
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "C35_1[0]", "")
                        End If
                    End If

                    'C36
                    'NGOHCS B2010.2
                    'pdfFormFields.SetField(pdfFieldPath & "C36[0]", nTotal + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW")))
                    If Not IsDBNull(dr2("TP_STATUS")) Then
                        If dr2("TP_STATUS") = "1" Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        ElseIf dr2("TP_STATUS") = "2" And CDbl(pdfForm.GetYA) >= 2007 And dr2("TP_TYPE_ASSESSMENT") = "3" Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        ElseIf dr2("TP_STATUS") = "2" And CDbl(pdfForm.GetYA) = 2006 And dr2("TP_TYPE_ASSESSMENT") = "1" And CDbl(dr("TC_INCOME_TRANSFER_FROM_HW")) = 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        ElseIf (dr2("TP_GENDER") = "1" And dr2("TP_STATUS") = "2" And dr2("TP_TYPE_ASSESSMENT") = "1" And _
                            dr2("TP_ASSESSMENTON") = "1") Or (dr2("TP_GENDER") = "2" And dr2("TP_STATUS") = "2" And dr2("TP_TYPE_ASSESSMENT") = "1" And dr2("TP_ASSESSMENTON") = "2") Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", CDbl(FormatNumber(((CDbl(dr("TC_TOTAL_INCOME_2")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))).ToString), 0)))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        End If
                    End If

                End If
                dr2.Close()
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try
    End Sub
    Private Sub Page5()
        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim ComputationKey As String
        Dim prmOledb(1) As SqlParameter
        pdfFieldPath = pdfSubFormName & "Page7[0]."
        pdfFormFields.SetField(pdfFieldPath & "Nama7[0]", RefName)
        pdfFormFields.SetField(pdfFieldPath & "Ruj7[0]", pdfForm.GetRefNo)
        Try
            'Part D
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            prmOledb(1) = New SqlParameter("@YA", pdfForm.GetYA)
            'weihong
            ds = datHandler.GetData("SELECT TC_CHARGEABLE_INCOME, TC_TOTAL_INCOME_TAX, TC_INCOME_TAX_CHARGED," _
                            & " TC_SEC110_DIVIDEND, TC_SEC110_OTHERS, TC_SEC130, TC_5, TC_TAX_PAYABLE, TC_TAX_REPAYMENT, TC_KEY, TC_2, TC_TAX_SCH1_INCOME, TC_TAX_SCH1_TAX" _
                            & " FROM TAX_COMPUTATION WHERE TC_REF_NO=@ref_no AND TC_YA=@ya", prmOledb(0), prmOledb(1))
            pdfFormFields.SetField(pdfFieldPath & "D1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(0).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "D3[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(1).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D5[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(2).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D6[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(3).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D7[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(4).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D8[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(10).ToString, 2).ToString.Replace(".", "")))
            'weihong
            'If Not IsDBNull(ds.Tables(0).Rows(0).Item(11)) Then
            '    pdfFormFields.SetField(pdfFieldPath & "D2_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(11).ToString, 0).ToString.Replace(".", "")))
            'End If
            'If Not IsDBNull(ds.Tables(0).Rows(0).Item(12)) Then
            '    pdfFormFields.SetField(pdfFieldPath & "D2_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(12).ToString, 2).ToString.Replace(".", "")))
            'End If
            'pdfFormFields.SetField(pdfFieldPath & "D8A[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(10).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D9[0]", FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(6).ToString), 2).ToString.Replace(".", "").Replace(",", ""))
            If (CDbl(ds.Tables(0).Rows(0).Item(6).ToString)) < CDbl(ds.Tables(0).Rows(0).Item(2).ToString) Then
                pdfFormFields.SetField(pdfFieldPath & "D10[0]", FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(2).ToString) - CDbl(ds.Tables(0).Rows(0).Item(6).ToString), 2).ToString.Replace(".", "").Replace(",", ""))
                pdfFormFields.SetField(pdfFieldPath & "D11[0]", FormatNumber(CDbl("0"), 2).ToString.Replace(".", ""))
            Else
                pdfFormFields.SetField(pdfFieldPath & "D10[0]", FormatNumber(CDbl("0"), 2).ToString.Replace(".", ""))
                pdfFormFields.SetField(pdfFieldPath & "D11[0]", FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(6).ToString) - CDbl(ds.Tables(0).Rows(0).Item(2).ToString), 2).ToString.Replace(".", "").Replace(",", ""))
            End If
            ComputationKey = ds.Tables(0).Rows(0).Item(9).ToString

            ds = datHandler.GetData("SELECT CHARGEABLE0, INCOME0, CHARGEABLE1, INCOME1, CHARGEABLE2, INCOME2," _
                            & " CHARGEABLE3, INCOME3, CHARGEABLE4, INCOME4, CHARGEABLE5, INCOME5, CHARGEABLE6, RATE1, INCOME6" _
                            & " FROM CHARGEABLE_INCOME WHERE TC_KEY=" + ComputationKey)
            pdfFormFields.SetField(pdfFieldPath & "D2a_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(0).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "D2a_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(1).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D2b_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(2).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "D2b_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(3).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D2c_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(4).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "D2c_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(5).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D2d_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(6).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "D2d_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(7).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D2e_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(8).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "D2e_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(9).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D2f_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(10).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "D2f_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(11).ToString, 2).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D2g_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(12).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "D2g_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(13).ToString, 1).ToString.Replace(".", "")))
            pdfFormFields.SetField(pdfFieldPath & "D2g_3[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(14).ToString, 2).ToString.Replace(".", "")))

            ds = datHandler.GetData("SELECT TCR_AMOUNT" _
                & " FROM TAX_REBATE WHERE TC_KEY=" + ComputationKey + " AND TCR_KEY=5")
            pdfFormFields.SetField(pdfFieldPath & "D4[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(0).ToString, 2).ToString.Replace(".", "")))

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try
    End Sub
    Private Sub Page6()
        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim i As Integer
        Dim prmOledb(1) As SqlParameter
        Dim strArray(1) As String
        pdfFieldPath = pdfSubFormName & "Page8[0]."
        pdfFormFields.SetField(pdfFieldPath & "Nama8[0]", RefName)
        pdfFormFields.SetField(pdfFieldPath & "Ruj8[0]", pdfForm.GetRefNo)
        Try
            'Part E
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            prmOledb(1) = New SqlParameter("@YA", pdfForm.GetYA)
            ds = datHandler.GetData("SELECT TC_TAX_PAYABLE, TC_INSTALLMENT_PAYMENT_SELF, TC_INSTALLMENT_PAYMENT_HW," _
                            & " TC_BALANCE_TAX_PAYABLE, TC_BALANCE_TAX_OVERPAID" _
                            & " FROM TAX_COMPUTATION WHERE TC_REF_NO=@ref_no AND TC_YA=@ya", prmOledb(0), prmOledb(1))

            If CDbl(ds.Tables(0).Rows(0).Item(0).ToString) > 0 Then
                pdfFormFields.SetField(pdfFieldPath & "E1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(0).ToString, 2).ToString.Replace(".", "")))
            Else
                pdfFormFields.SetField(pdfFieldPath & "E1[0]", "000")
            End If
            If CDbl(ds.Tables(0).Rows(0).Item(1).ToString) + CDbl(ds.Tables(0).Rows(0).Item(2).ToString) > 0 Then
                pdfFormFields.SetField(pdfFieldPath & "E2[0]", CDbl(FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(1)) + CDbl(ds.Tables(0).Rows(0).Item(2)), 2).ToString.Replace(".", "")))
            Else
                pdfFormFields.SetField(pdfFieldPath & "E2[0]", "000")
            End If
            If CDbl(ds.Tables(0).Rows(0).Item(3).ToString) Then
                pdfFormFields.SetField(pdfFieldPath & "E3[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(3).ToString, 2).ToString.Replace(".", "")))
            Else
                pdfFormFields.SetField(pdfFieldPath & "E3[0]", "000")
            End If
            If CDbl(ds.Tables(0).Rows(0).Item(4).ToString) Then
                pdfFormFields.SetField(pdfFieldPath & "E4[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(4).ToString, 2).ToString.Replace(".", "")))
            Else
                pdfFormFields.SetField(pdfFieldPath & "E4[0]", "000")
            End If

            'Part F
            dr = datHandler.GetDataReader("SELECT PY_KEY " _
                            & "FROM PRECEDING_YEAR WHERE PY_REF_NO ='" + pdfForm.GetRefNo + "' AND PY_YA ='" + pdfForm.GetYA + "'")
            If dr.Read Then
                i = 0
                dr1 = datHandler.GetDataReader("SELECT TOP 3 *" _
                                & " FROM PRECEDING_YEAR_DETAIL WHERE PY_KEY=" + dr("PY_KEY").ToString)
                While dr1.Read
                    i = i + 1
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_1[0]", dr1("PY_INCOME_TYPE").ToString.ToUpper)
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_2[0]", dr1("PY_PAYMENT_YEAR").ToString)
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_3[0]", CDbl(FormatNumber(dr1("PY_AMOUNT"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_4[0]", CDbl(FormatNumber(dr1("PY_EPF"), 0)))
                End While
                While i < 3
                    i = i + 1
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_3[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "F" + CStr(i) + "_4[0]", 0)
                End While
                dr1.Close()
            Else
                pdfFormFields.SetField(pdfFieldPath & "F1_3[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "F1_4[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "F2_3[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "F2_4[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "F3_3[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "F3_4[0]", 0)
            End If
            dr.Close()

            'Part G
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            ds = datHandler.GetData("SELECT TP_ADM_NAME, (TP_ADM_IC_NEW1 + TP_ADM_IC_NEW2 + TP_ADM_IC_NEW3)," _
                        & " TP_ADM_IC_OLD, TP_ADM_POLICE_NO, TP_ADM_ARMY_NO, TP_ADM_PASSPORT_NO" _
                        & " FROM TAXP_PROFILE WHERE (TP_REF_NO1 + TP_REF_NO2 + TP_REF_NO3)=@ref_no", prmOledb(0))
            If ds.Tables(0).Rows(0).Item(0) <> "" Then
                strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString.ToUpper, 28)
                pdfFormFields.SetField(pdfFieldPath & "G1_1[0]", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "G1_2[0]", strArray(1))
            Else
                pdfFormFields.SetField(pdfFieldPath & "G1_1[0]", "---")
                pdfFormFields.SetField(pdfFieldPath & "G1_2[0]", "---")
            End If
            If ds.Tables(0).Rows(0).Item(1) <> "" Then
                pdfFormFields.SetField(pdfFieldPath & "G2[0]", ds.Tables(0).Rows(0).Item(1).ToString)
            Else
                pdfFormFields.SetField(pdfFieldPath & "G2[0]", "---")
            End If
            If ds.Tables(0).Rows(0).Item(2) <> "" Then
                pdfFormFields.SetField(pdfFieldPath & "G3[0]", ds.Tables(0).Rows(0).Item(2).ToString)
            Else
                pdfFormFields.SetField(pdfFieldPath & "G3[0]", "---")
            End If
            If ds.Tables(0).Rows(0).Item(3) <> "" Then
                pdfFormFields.SetField(pdfFieldPath & "G4[0]", ds.Tables(0).Rows(0).Item(3).ToString)
            Else
                pdfFormFields.SetField(pdfFieldPath & "G4[0]", "---")
            End If
            If ds.Tables(0).Rows(0).Item(4) <> "" Then
                pdfFormFields.SetField(pdfFieldPath & "G5[0]", ds.Tables(0).Rows(0).Item(4).ToString)
            Else
                pdfFormFields.SetField(pdfFieldPath & "G5[0]", "---")
            End If
            If ds.Tables(0).Rows(0).Item(5) <> "" Then
                pdfFormFields.SetField(pdfFieldPath & "G6[0]", ds.Tables(0).Rows(0).Item(5).ToString)
            Else
                pdfFormFields.SetField(pdfFieldPath & "G6[0]", "---")
            End If

            'Part H
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            prmOledb(1) = New SqlParameter("@YA", pdfForm.GetYA)
            ds = datHandler.GetData("SELECT TC_AL_CY_UNASORBED_LOSS, TC_AL_BAL_UNASORBED_LOSS, TC_AL_BALANCE_CF," _
                            & " TC_PIONEER, TC_PIONEER_CF" _
                            & " FROM TAX_COMPUTATION WHERE TC_REF_NO=@ref_no AND TC_YA=@ya", prmOledb(0), prmOledb(1))
            pdfFormFields.SetField(pdfFieldPath & "H1a_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(0).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "H1b[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(1).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "H1c[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(2).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "H1d_1[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(3).ToString, 0)))
            pdfFormFields.SetField(pdfFieldPath & "H1d_2[0]", CDbl(FormatNumber(ds.Tables(0).Rows(0).Item(4).ToString, 0)))


            'NGOHCS B2010.2 
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            prmOledb(1) = New SqlParameter("@YA", pdfForm.GetYA)
            ds = datHandler.GetData("select SUM(cast(tca_cbl as money)) from tax_adjusted_loss where tc_key in " _
                            & "(select tc_key from tax_computation where tc_ref_no =@ref_no and tc_ya=@ya)", prmOledb(0), prmOledb(1))

            pdfFormFields.SetField(pdfFieldPath & "H1a_1[0]", FormatNumber(CDbl("0"), 0))
            If ds.Tables.Count > 0 Then
                If ds.Tables(0).Rows.Count > 0 Then
                    If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "H1a_1[0]", FormatNumber(CDbl(ds.Tables(0).Rows(0).Item(0).ToString), 0).Replace(",", ""))
                    End If
                End If
            End If
            ds.Dispose()
            'NGOHCS B2010.2 END

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try
    End Sub
    Private Sub Page7()
        Dim pdfFieldPath As String
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim i As Integer
        Dim nTotal As Double, nTotal2 As Double
        Dim boolHasRecord As Boolean = True
        pdfFieldPath = pdfSubFormName & "Page9[0]."
        pdfFormFields.SetField(pdfFieldPath & "Nama9[0]", RefName)
        pdfFormFields.SetField(pdfFieldPath & "Ruj9[0]", pdfForm.GetRefNo)
        Try
            'Part H
            nTotal = 0
            nTotal2 = 0
            i = 97 'ASCII code for "a"
            dr = datHandler.GetDataReader("SELECT TOP 2 ADCA_UTIL, ADCA_BAL_CF" _
                            & " FROM INCOME_ADJUSTED WHERE ADJ_REF_NO='" + pdfForm.GetRefNo + "' AND ADJ_YA ='" + pdfForm.GetYA + "' ORDER BY ADJ_BUSINESS")
            pdfFormFields.SetField(pdfFieldPath & "H2a_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H2a_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H2b_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H2b_2[0]", 0)
            'If dr.Read Then
            While dr.Read
                pdfFormFields.SetField(pdfFieldPath & "H2" + Chr(i) + "_1[0]", CDbl(FormatNumber(dr("ADCA_UTIL"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "H2" + Chr(i) + "_2[0]", CDbl(FormatNumber(dr("ADCA_BAL_CF"), 0)))
                nTotal = nTotal + CDbl(FormatNumber(dr("ADCA_UTIL"), 0))
                nTotal2 = nTotal2 + CDbl(FormatNumber(dr("ADCA_BAL_CF"), 0))
                i = i + 1
            End While
            'Else
            'End If
            dr.Close()
            dr = datHandler.GetDataReader("SELECT SUM(cast(ADCA_UTIL as money)) AS TEMPNUM1, SUM(cast(ADCA_BAL_CF as money)) AS TEMPNUM2" _
                            & " FROM INCOME_ADJUSTED WHERE ADJ_REF_NO='" + pdfForm.GetRefNo + "' AND ADJ_YA ='" + pdfForm.GetYA + "'")
            If dr.Read Then
                If Not IsDBNull(dr("TEMPNUM1")) And Not IsDBNull(dr("TEMPNUM2")) Then
                    nTotal = CDbl(dr("TEMPNUM1")) - nTotal
                    nTotal2 = CDbl(dr("TEMPNUM2")) - nTotal2
                    If nTotal >= 0 Or nTotal2 >= 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "H2c_1[0]", nTotal)
                        pdfFormFields.SetField(pdfFieldPath & "H2c_2[0]", nTotal2)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "H2c_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "H2c_2[0]", 0)
                End If
            Else
                pdfFormFields.SetField(pdfFieldPath & "H2c_1[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "H2c_2[0]", 0)
            End If
            dr.Close()

            nTotal = 0
            nTotal2 = 0
            i = 100 'ASCII code for "d"
            dr = datHandler.GetDataReader("SELECT TOP 2 PSCA_UTIL, PSCA_BAL_CF" _
                            & " FROM INCOME_PARTNERSHIP WHERE PN_REF_NO='" + pdfForm.GetRefNo + "' AND PN_YA ='" + pdfForm.GetYA + "' ORDER BY PS_SOURCE")
            pdfFormFields.SetField(pdfFieldPath & "H2d_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H2d_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H2e_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H2e_2[0]", 0)
            'If dr.Read Then
            While dr.Read
                pdfFormFields.SetField(pdfFieldPath & "H2" + Chr(i) + "_1[0]", CDbl(FormatNumber(dr("PSCA_UTIL"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "H2" + Chr(i) + "_2[0]", CDbl(FormatNumber(dr("PSCA_BAL_CF"), 0)))
                nTotal = nTotal + CDbl(FormatNumber(dr("PSCA_UTIL"), 0))
                nTotal2 = nTotal2 + CDbl(FormatNumber(dr("PSCA_BAL_CF"), 0))
                i = i + 1
            End While
            'Else
            'End If
            dr.Close()
            dr = datHandler.GetDataReader("SELECT SUM(cast(PSCA_UTIL as money)) AS TEMPNUM1, SUM(cast(PSCA_BAL_CF as money)) AS TEMPNUM2" _
                            & " FROM INCOME_PARTNERSHIP WHERE PN_REF_NO='" + pdfForm.GetRefNo + "' AND PN_YA ='" + pdfForm.GetYA + "'")
            If dr.Read Then
                If Not IsDBNull(dr("TEMPNUM1")) And Not IsDBNull(dr("TEMPNUM2")) Then
                    nTotal = CDbl(dr("TEMPNUM1")) - nTotal
                    nTotal2 = CDbl(dr("TEMPNUM2")) - nTotal2
                    If nTotal >= 0 Or nTotal2 >= 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "H2f_1[0]", nTotal)
                        pdfFormFields.SetField(pdfFieldPath & "H2f_2[0]", nTotal2)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "H2f_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "H2f_2[0]", 0)
                End If
            Else
                pdfFormFields.SetField(pdfFieldPath & "H2f_1[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "H2f_2[0]", 0)
            End If
            dr.Close()

            dr = datHandler.GetDataReader("SELECT NR_SECTION, NR_GROSS_TOTAL, NR_WITHHOLD, NR_WITHHOLD_107A" _
                            & " FROM NON_RESIDENT WHERE NR_REF_NO='" + pdfForm.GetRefNo + "' AND NR_YA ='" + pdfForm.GetYA + "' ORDER BY NR_SECTION")
            pdfFormFields.SetField(pdfFieldPath & "H3a_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H3a_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H3b_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H3b_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H3c_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H3c_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H3d_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H3d_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H3e_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "H3e_2[0]", 0)
            While dr.Read
                If dr("NR_SECTION") = 1 Then
                    pdfFormFields.SetField(pdfFieldPath & "H3a_1[0]", CDbl(FormatNumber(dr("NR_GROSS_TOTAL"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "H3a_2[0]", CDbl(FormatNumber(CDbl(dr("NR_WITHHOLD")) + CDbl(dr("NR_WITHHOLD_107A")), 0)))
                ElseIf dr("NR_SECTION") = 2 Then
                    pdfFormFields.SetField(pdfFieldPath & "H3b_1[0]", CDbl(FormatNumber(dr("NR_GROSS_TOTAL"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "H3b_2[0]", CDbl(FormatNumber(dr("NR_WITHHOLD"), 0)))
                ElseIf dr("NR_SECTION") = 3 Then
                    pdfFormFields.SetField(pdfFieldPath & "H3c_1[0]", CDbl(FormatNumber(dr("NR_GROSS_TOTAL"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "H3c_2[0]", CDbl(FormatNumber(dr("NR_WITHHOLD"), 0)))
                ElseIf dr("NR_SECTION") = 4 Then
                    pdfFormFields.SetField(pdfFieldPath & "H3d_1[0]", CDbl(FormatNumber(dr("NR_GROSS_TOTAL"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "H3d_2[0]", CDbl(FormatNumber(dr("NR_WITHHOLD"), 0)))
                ElseIf dr("NR_SECTION") = 6 Then
                    'NGOHCS B2010.2
                    pdfFormFields.SetField(pdfFieldPath & "H3e_1[0]", CDbl(FormatNumber(dr("NR_GROSS_TOTAL"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "H3e_2[0]", CDbl(FormatNumber(dr("NR_WITHHOLD"), 0)))
                    'NGOHCS B2010.2 END
                End If

            End While
            dr.Close()

            'Part J
            dr = datHandler.GetDataReader("SELECT ADJ_KEY " _
                            & "FROM INCOME_ADJUSTED WHERE ADJ_REF_NO ='" + pdfForm.GetRefNo + "' AND ADJ_YA ='" + pdfForm.GetYA + "'")
            'NGOHCS B+ C2009.1 (SU11)
            'If dr.Read Then
            i = 0
            nTotal = 0
            While dr.Read
                dr1 = datHandler.GetDataReader("SELECT ADJD_CLAIM_CODE, ADJD_AMOUNT" _
                                                & " FROM INCOME_ADJ_FURTHER WHERE ADJ_KEY=" + dr("ADJ_KEY").ToString + " Order By ADJD_ID, ADJD_NO")
                While dr1.Read
                    boolHasRecord = True
                    i = i + 1
                    If i <= 4 Then
                        pdfFormFields.SetField(pdfFieldPath & "J" + CStr(i) + "_1[0]", dr1("ADJD_CLAIM_CODE").ToString)
                        pdfFormFields.SetField(pdfFieldPath & "J" + CStr(i) + "_2[0]", CDbl(FormatNumber(dr1("ADJD_AMOUNT").ToString, 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1("ADJD_AMOUNT"), 0))
                    End If
                End While
                dr1.Close()
            End While
            While (i < 4)
                i = i + 1
                pdfFormFields.SetField(pdfFieldPath & "J" + CStr(i) + "_2[0]", 0)
            End While
            'Else
            If Not boolHasRecord Then
                pdfFormFields.SetField(pdfFieldPath & "J1_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "J2_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "J3_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "J4_2[0]", 0)
            End If
            'End If
            'NGOHCS B+ C2009.1 (SU11) END
            dr.Close()
            pdfFormFields.SetField(pdfFieldPath & "J5[0]", nTotal)

            'Part K
            dr = datHandler.GetDataReader("SELECT TC_KEY " _
                & "FROM TAX_COMPUTATION WHERE TC_REF_NO ='" + pdfForm.GetRefNo + "' AND TC_YA ='" + pdfForm.GetYA + "'")
            If dr.Read Then
                dr1 = datHandler.GetDataReader("SELECT TIC_KEY, TIC_CF" _
                                           & " FROM TAX_INCENTIVE_CLAIM WHERE TC_KEY=" + dr("TC_KEY").ToString)
                While dr1.Read
                    If dr1("TIC_KEY") = 3 Then
                        If dr1("TIC_CF") <> "" Then
                            pdfFormFields.SetField(pdfFieldPath & "K1[0]", CDbl(FormatNumber(dr1("TIC_CF"), 0)))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "K1[0]", 0)
                        End If
                    ElseIf dr1("TIC_KEY") = 5 Then
                        If dr1("TIC_CF") <> "" Then
                            pdfFormFields.SetField(pdfFieldPath & "K2[0]", CDbl(FormatNumber(dr1("TIC_CF"), 0)))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "K2[0]", 0)
                        End If
                    End If
                End While
                dr1.Close()
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try
    End Sub
    Private Sub Page8()
        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing, dr1 As SqlDataReader = Nothing
        Dim prmOledb(1) As SqlParameter
        Dim strArray(1) As String
        Dim cSQL As String
        Dim PLKey As Long
        Dim nTotal As Double, nTotal2 As Double, nTotal3 As Double
        pdfFieldPath = pdfSubFormName & "Page10[0]."
        pdfFormFields.SetField(pdfFieldPath & "Nama10[0]", RefName)
        pdfFormFields.SetField(pdfFieldPath & "Ruj10[0]", pdfForm.GetRefNo)
        Try
            'Part L Profit and Loss Account
            dr = datHandler.GetDataReader("SELECT * FROM PROFIT_LOSS_ACCOUNT WHERE PL_REF_NO = '" & pdfForm.GetRefNo & "' AND PL_YA = '" & pdfForm.GetYA & "' and PL_MAINCOMPANY = '1' order by PL_KEY")
            If dr.Read Then
                cSQL = "SELECT PL_MAIN_BUSINESS, PL_KEY, PL_COMPANY" _
                & " FROM PROFIT_LOSS_ACCOUNT WHERE PL_REF_NO = '" & pdfForm.GetRefNo & "' AND PL_YA = '" & pdfForm.GetYA & "' and PL_MAINCOMPANY = '1' ORDER BY PL_KEY "
            Else
                cSQL = "SELECT PL_MAIN_BUSINESS, PL_KEY, PL_COMPANY" _
                & " FROM PROFIT_LOSS_ACCOUNT WHERE PL_REF_NO = '" & pdfForm.GetRefNo & "' AND PL_YA = '" & pdfForm.GetYA & "' ORDER BY PL_KEY "
            End If
            dr.Close()
            PnLMainBusiness = ""
            dr = datHandler.GetDataReader(cSQL)
            If dr.Read Then
                PnLMainBusiness = dr("PL_MAIN_BUSINESS")
                dr1 = datHandler.GetDataReader("SELECT BC_BUS_ENTITY, BC_CODE FROM BUSINESS_SOURCE WHERE BC_KEY = '" & pdfForm.GetRefNo & "' AND BC_YA = '" & pdfForm.GetYA & "' AND BC_BUSINESSSOURCE = '" & Trim(dr("PL_MAIN_BUSINESS")) & "'")
                If dr1.Read Then
                    strArray = SplitText(dr1("BC_BUS_ENTITY").ToString.ToUpper, 28)
                    pdfFormFields.SetField(pdfFieldPath & "L1_1[0]", strArray(0))
                    pdfFormFields.SetField(pdfFieldPath & "L1_2[0]", strArray(1))
                    pdfFormFields.SetField(pdfFieldPath & "L1A[0]", dr1("BC_CODE"))
                End If
                dr1.Close()
                PLKey = CLng(dr("PL_KEY"))
                nTotal2 = 0
                nTotal3 = 0
                'Sales
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(PL_AMOUNT as money)) FROM PL_SALES WHERE PL_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L2[0]", CDbl(FormatNumber(dr1(0), 0)).ToString)
                        nTotal3 = CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L2[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L2[0]", "0")
                End If
                dr1.Close()
                'Opening Stock
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(PL_AMOUNT as money)) FROM PL_OPENSTOCK WHERE PL_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L3[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal2 = CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L3[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L3[0]", "0")
                End If
                dr1.Close()
                'Purchase and Cost of Production
                nTotal = 0
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(PL_AMOUNT as money)) FROM PL_PURCHASE WHERE PL_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_PRODUCTION_COST WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                nTotal2 = nTotal2 + nTotal
                pdfFormFields.SetField(pdfFieldPath & "L4[0]", nTotal)
                'Closing Stock
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(PL_AMOUNT as money)) FROM PL_CLOSESTOCK WHERE PL_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L5[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal2 = nTotal2 - CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L5[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L5[0]", "0")
                End If
                dr1.Close()
                'Cost of Sales
                pdfFormFields.SetField(pdfFieldPath & "L6[0]", nTotal2)
                'Gross Profit / Loss
                nTotal3 = nTotal3 - nTotal2
                If nTotal3 < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L7_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L7_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L7_2[0]", Abs(nTotal3))
                nTotal = 0
                nTotal2 = 0
                'Other Business Income
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_OTHERBUSINESS WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = CDbl(FormatNumber(dr1(0), 0))
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L8[0]", "0")
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT [PL_KEY] FROM [PROFIT_LOSS_ACCOUNT]" _
                        & " WHERE [PL_REF_NO] = '" & pdfForm.GetRefNo & "' AND [PL_YA] = '" & pdfForm.GetYA & "' and [PL_KEY] <> " & CStr(PLKey))
                While dr1.Read()
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = nTotal2 + OtherSource_GrossProfitLoss(CLng(dr1(0)), dr("PL_COMPANY").ToString, PnLMainBusiness, pdfForm.GetRefNo, pdfForm.GetYA)
                    End If
                End While
                dr1.Close()
                pdfFormFields.SetField(pdfFieldPath & "L8[0]", nTotal2)
                nTotal = nTotal2

                'Dividends
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_NONBUSINESS WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 47")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L9[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L9[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L9[0]", "0")
                End If
                dr1.Close()
                'Interest and discounts
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_NONBUSINESS WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 50")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L10[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L10[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L10[0]", "0")
                End If
                dr1.Close()
                'Rents, royalties and premiums
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_NONBUSINESS WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE BETWEEN 48 and 49")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L11[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L11[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L11[0]", "0")
                End If
                dr1.Close()
                nTotal2 = 0
                'Other Income
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_NONBUSINESS WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 51")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_INCOME_NONTAXABLE WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = nTotal2 + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                pdfFormFields.SetField(pdfFieldPath & "L12[0]", nTotal2)
                nTotal = nTotal + nTotal2
                pdfFormFields.SetField(pdfFieldPath & "L13[0]", nTotal)
                nTotal3 = nTotal3 + nTotal
                nTotal = 0
                'Loan interest
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 11")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L14[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L14[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L14[0]", "0")
                End If
                dr1.Close()
                'Salaries and wages
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 12")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L15[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L15[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L15[0]", "0")
                End If
                dr1.Close()
                'Rental / Lease
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 13")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L16[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L16[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L16[0]", "0")
                End If
                dr1.Close()
                'Contracts and subcontracts
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 14")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L17[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L17[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L17[0]", "0")
                End If
                dr1.Close()
                'Commissions
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 15")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L18[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L18[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L18[0]", "0")
                End If
                dr1.Close()
                'Bad debts
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 16")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L19[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L19[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L19[0]", "0")
                End If
                dr1.Close()
                'Travelling and transport
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 17")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L20[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L20[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L20[0]", "0")
                End If
                dr1.Close()
                'Repair and maintenance
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 52")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L21[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L21[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L21[0]", "0")
                End If
                dr1.Close()
                'Promotion and advertisement
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND EXA_PLTYPE = 53")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "L22[0]", CDbl(FormatNumber(dr1(0), 0)))
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L22[0]", "0")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L22[0]", "0")
                End If
                dr1.Close()
                nTotal2 = 0
                'Other expenses
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & "AND (EXA_PLTYPE between 18 and 20 or EXA_PLTYPE = 46)")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXP_NONALLOWLOSS WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = nTotal2 + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXP_NONALLOWEXPEND WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = nTotal2 + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXP_PERSONAL WHERE EXA_KEY = " & CStr(PLKey))
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal2 = nTotal2 + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                pdfFormFields.SetField(pdfFieldPath & "L23[0]", nTotal2)
                nTotal = nTotal + nTotal2
                pdfFormFields.SetField(pdfFieldPath & "L24[0]", nTotal)
                nTotal3 = nTotal3 - nTotal
                If nTotal3 < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L25_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L25_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L25_2[0]", Abs(nTotal3))
                nTotal = 0
                'Non-allowable expenses
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXPENSES WHERE EXA_KEY = " & CStr(PLKey) & " AND [EXA_DEDUCTIBLE]='No'")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXP_NONALLOWEXPEND WHERE EXA_KEY = " & CStr(PLKey) & " AND [EXA_DEDUCTIBLE]='No'")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_EXP_PERSONAL WHERE EXA_KEY = " & CStr(PLKey) & " AND [EXA_DEDUCTIBLE]='No'")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                dr1 = datHandler.GetDataReader("SELECT SUM(cast(EXA_AMOUNT as money)) FROM PL_PRODUCTION_COST WHERE EXA_KEY = " & CStr(PLKey) & "AND [EXA_DEDUCTIBLE]='No' and (EXA_PLTYPE = 43 or EXA_PLTYPE = 45)")
                If dr1.Read Then
                    If Not IsDBNull(dr1(0)) Then
                        nTotal = nTotal + CDbl(FormatNumber(dr1(0), 0))
                    End If
                End If
                dr1.Close()
                pdfFormFields.SetField(pdfFieldPath & "L26[0]", nTotal)
            Else
                pdfFormFields.SetField(pdfFieldPath & "L1A[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L2[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L3[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L4[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L5[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L6[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L7_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L7_2[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L8[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L9[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L10[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L11[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L12[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L13[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L14[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L15[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L16[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L17[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L18[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L19[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L20[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L21[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L22[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L23[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L24[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L25_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L25_2[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "L26[0]", "0")
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try
    End Sub
    Private Sub Page9()
        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        pdfFieldPath = pdfSubFormName & "Page11[0]."
        pdfFormFields.SetField(pdfFieldPath & "Nama11[0]", RefName)
        pdfFormFields.SetField(pdfFieldPath & "Ruj11[0]", pdfForm.GetRefNo)
        Try
            'Part L Balance Sheet
            dr = datHandler.GetDataReader("SELECT *" _
                & " FROM BALANCE_SHEET WHERE BS_REF_NO = '" & pdfForm.GetRefNo & "' AND BS_YA = '" & pdfForm.GetYA & "' AND [BS_SOURCENO] = '" & Trim(PnLMainBusiness) + "' ORDER BY BS_SOURCENO")
            If dr.Read Then
                pdfFormFields.SetField(pdfFieldPath & "L27[0]", CDbl(FormatNumber(dr("BS_LAND"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L28[0]", CDbl(FormatNumber(dr("BS_MACHINERY"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L29[0]", CDbl(FormatNumber(dr("BS_TRANSPORT"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L30[0]", CDbl(FormatNumber(dr("BS_OTH_FA"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L31[0]", CDbl(FormatNumber(dr("BS_TOT_FA"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L32[0]", CDbl(FormatNumber(dr("BS_INVESTMENT"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L33[0]", CDbl(FormatNumber(dr("BS_STOCK"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L34[0]", CDbl(FormatNumber(dr("BS_TRADE_DEBTORS"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L35[0]", CDbl(FormatNumber(dr("BS_OTH_DEBTORS"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L36[0]", CDbl(FormatNumber(dr("BS_CASH"), 0)))
                If CDbl(FormatNumber(dr("BS_BANK"), 0)) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L37_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L37_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L37_2[0]", Abs(CDbl(FormatNumber(dr("BS_BANK"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "L38[0]", CDbl(FormatNumber(dr("BS_OTH_CA"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L39[0]", CDbl(FormatNumber(dr("BS_TOT_CA"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L40[0]", CDbl(FormatNumber(dr("BS_TOT_ASSETS"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L41[0]", CDbl(FormatNumber(dr("BS_LOAN"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L42[0]", CDbl(FormatNumber(dr("BS_TRADE_CR"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L43[0]", (CDbl(FormatNumber(dr("BS_OTHER_CR"), 0)) + CDbl(FormatNumber(dr("BS_OTH_LIAB"), 0)) + CDbl(FormatNumber(dr("BS_LT_LIAB"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "L44[0]", CDbl(FormatNumber(dr("BS_TOT_LIAB"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "L45[0]", CDbl(FormatNumber(dr("BS_CAPITALACCOUNT"), 0)))
                If CDbl(FormatNumber(dr("BS_BROUGHT_FORWARD"), 0)) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L46_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L46_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L46_2[0]", Abs(CDbl(FormatNumber(dr("BS_BROUGHT_FORWARD"), 0))))
                If CDbl(FormatNumber(dr("BS_CY_PROFITLOSS"), 0)) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L47_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L47_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L47_2[0]", Abs(CDbl(FormatNumber(dr("BS_CY_PROFITLOSS"), 0))))
                If (CDbl(FormatNumber(dr("BS_CAP_CONTRIBUTION"), 0)) - CDbl(FormatNumber(dr("BS_DRAWING"), 0))) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L48_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L48_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L48_2[0]", Abs(CDbl(FormatNumber(dr("BS_CAP_CONTRIBUTION"), 0)) - CDbl(FormatNumber(dr("BS_DRAWING"), 0))))
                If CDbl(FormatNumber(dr("BS_CARRIED_FORWARD"), 0)) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "L49_1[0]", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L49_1[0]", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L49_2[0]", Abs(CDbl(FormatNumber(dr("BS_CARRIED_FORWARD"), 0))))
            Else
                pdfFormFields.SetField(pdfFieldPath & "L27[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L28[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L29[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L30[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L31[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L32[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L33[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L34[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L35[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L36[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L37_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L37_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L38[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L39[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L40[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L41[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L42[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L43[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L44[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L45[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L46_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L46_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L47_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L47_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L48_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L48_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "L49_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "L49_2[0]", 0)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try
    End Sub
    Private Sub Page10()
        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim strArray(1) As String
        Dim strLine As String = ""
        pdfFieldPath = pdfSubFormName & "Page12[0]."
        pdfFormFields.SetField(pdfFieldPath & "Nama12[0]", RefName)
        pdfFormFields.SetField(pdfFieldPath & "Ruj12[0]", pdfForm.GetRefNo)
        Try
            'Declaration
            If pdfForm.GetDeclarationReturn = 1 Then
                dr = datHandler.GetDataReader("SELECT * FROM [TAXP_PROFILE] WHERE [TP_REF_NO1] & [TP_REF_NO2] & [TP_REF_NO3] = '" & pdfForm.GetRefNo & "'")
                If dr.Read() Then
                    strArray = SplitText(RefName, 28)
                    pdfFormFields.SetField(pdfFieldPath & "Akuan1_1[0]", strArray(0))
                    pdfFormFields.SetField(pdfFieldPath & "Akuan1_2[0]", strArray(1))
                    If Len(Trim(dr("TP_IC_NEW_1") + Trim(dr("TP_IC_NEW_2")) + Trim(dr("TP_IC_NEW_3")))) > 0 Then
                        strLine = Trim(dr("TP_IC_NEW_1")) + Trim(dr("TP_IC_NEW_2")) + Trim(dr("TP_IC_NEW_3"))
                    ElseIf Len(Trim(dr("TP_PASSPORT_NO"))) > 0 Then
                        strLine = (dr("TP_PASSPORT_NO"))
                    ElseIf Len(Trim(dr("TP_POLICE_NO"))) > 0 Then
                        strLine = (dr("TP_POLICE_NO"))
                    ElseIf Len(Trim(dr("TP_ARMY_NO"))) > 0 Then
                        strLine = (dr("TP_ARMY_NO"))
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "Akuan2", strLine)
                    pdfFormFields.SetField(pdfFieldPath & "Akuan3", pdfForm.GetDeclarationReturn)
                    pdfFormFields.SetField(pdfFieldPath & "Akuan4", "")
                    pdfFormFields.SetField(pdfFieldPath & "Akuan5", pdfForm.GetDeclarationDate)
                End If
                dr.Close()
            ElseIf pdfForm.GetDeclarationReturn = 2 Then
                strArray = SplitText(pdfForm.GetDeclarationBy.ToString.ToUpper, 28)
                pdfFormFields.SetField(pdfFieldPath & "Akuan1_1[0]", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "Akuan1_2[0]", strArray(1))
                pdfFormFields.SetField(pdfFieldPath & "Akuan2", pdfForm.GetDeclarationID)
                pdfFormFields.SetField(pdfFieldPath & "Akuan3", pdfForm.GetDeclarationReturn)
                pdfFormFields.SetField(pdfFieldPath & "Akuan4", RefName)
                pdfFormFields.SetField(pdfFieldPath & "Akuan5", pdfForm.GetDeclarationDate)
            End If

            'Tax Agent
            ReDim strArray(1)
            dr = datHandler.GetDataReader("SELECT * FROM [TAXA_PROFILE] Where [TA_KEY] =" & pdfForm.GetTaxAgent)
            If dr.Read() Then
                If Not IsDBNull(dr("TA_CO_NAME")) Then
                    If Not String.IsNullOrEmpty(dr("TA_CO_NAME")) Then
                        strArray = SplitText(dr("TA_CO_NAME").ToString, 26)
                        pdfFormFields.SetField(pdfFieldPath & "NyataA_1", strArray(0))
                        pdfFormFields.SetField(pdfFieldPath & "NyataA_2", strArray(1))
                    End If
                End If
                pdfFormFields.SetField(pdfFieldPath & "Nyatab", FormatPhoneNumber("", dr("TA_TEL_NO").ToString, "", dr("TA_MOBILE").ToString))
                If Not IsDBNull(dr("TA_LICENSE")) Then
                    pdfFormFields.SetField(pdfFieldPath & "Nyatac", dr("TA_LICENSE"))
                End If
                pdfFormFields.SetField(pdfFieldPath & "NyataTarikh", FormatDate(Now))
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try
    End Sub
    Private Sub Page11()
        Dim pdfFieldPath As String
        pdfFieldPath = pdfSubFormName & "Page13[0]."
        pdfFormFields.SetField(pdfFieldPath & "Nama13[0]", RefName)
        pdfFormFields.SetField(pdfFieldPath & "Ruj13[0]", pdfForm.GetRefNo)

    End Sub
    Private Sub Page12()
        Dim pdfFieldPath As String = ""
        Dim strTempString As String = ""
        Dim dr As SqlDataReader = Nothing
        pdfFieldPath = pdfSubFormName & "Page14[0]."
        pdfFormFields.SetField(pdfFieldPath & "Nama14[0]", RefName)
        pdfFormFields.SetField(pdfFieldPath & "Ruj14[0]", pdfForm.GetRefNo)

        Try
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
                    strTempString = CStr(CDbl((FormatNumber(dr("TC_BALANCE_TAX_PAYABLE").ToString, 2))))
                    If CDbl((FormatNumber(dr("TC_BALANCE_TAX_PAYABLE").ToString, 2))) > 0 Then
                        strTempString = strTempString.ToString.Replace(".", "")
                    Else
                        strTempString = "000"
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "Slip2", strTempString)
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
                Case "Page2[0].A5[0]", _
                     "Page2[0].VII[0]", _
                     "Page2[0].VIII[0]", _
                     "Page2[0].A4[0]", _
                     "Page3[0].C3[0]", _
                     "Page3[0].G1_1[0]", _
                     "Page3[0].G1_2[0]", _
                     "Page3[0].G2_1[0]", _
                     "Page3[0].G2_2[0]"
                    pdfFormFields.SetField(de.Key.ToString, RTrim("---"))
            End Select
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
            If strMobilePrefix.Length = 2 Then
                strMobilePrefix = " " & strMobilePrefix
            End If
            strTemp = strMobilePrefix & strMobile
            strTemp = strTemp.Replace("-", "")
        End If
        Return strTemp

    End Function

    Protected Function OtherSource_GrossProfitLoss(ByVal cPNL_Key As Long, ByVal PNLCompany As String, ByVal PnLMainBusiness As String, ByVal strRefNo As String, ByVal strYA As String) As Double

        Dim i As Integer, J As Integer
        Dim arrSource() As String = Nothing
        Dim arrPNL(,) As Double
        Dim osTotal As Double
        Dim dr As SqlDataReader = Nothing

        i = 0
        osTotal = 0

        dr = datHandler.GetDataReader("SELECT [BC_BUSINESSSOURCE] FROM [BUSINESS_SOURCE] WHERE [BC_KEY] = '" & strRefNo & "'" _
                        & " AND [BC_YA] = '" & strYA & "' AND [BC_COMPANY] <> '" & Trim(PNLCompany) & "'")

        While dr.Read()
            i = i + 1
            ReDim Preserve arrSource(i)
            arrSource(i) = dr("BC_BUSINESSSOURCE")

        End While
        dr.Close()

        If i = 0 Then GoTo eSub

        ReDim arrPNL(i, 5)

        For J = 1 To UBound(arrPNL)

            arrPNL(J, 0) = 0 ''Sales
            arrPNL(J, 1) = 0 ''Opening Stock
            arrPNL(J, 2) = 0 ''Purchase
            arrPNL(J, 3) = 0 ''Cost of Production
            arrPNL(J, 4) = 0 ''Closing Stock
            arrPNL(J, 5) = 0 ''Gross Profit and Loss

            ''*** Sales
            dr = datHandler.GetDataReader("SELECT SUM(cast([PL_AMOUNT] as money)) FROM [PL_SALES] WHERE [PL_KEY] = " & cPNL_Key & " AND [PL_SOURCENO] = '" & Trim(arrSource(J)) & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then arrPNL(J, 0) = CDbl(dr.Item(0))
            End If
            dr.Close()
            ''*** Opening Stock
            dr = datHandler.GetDataReader("SELECT SUM(cast([PL_AMOUNT] as money)) FROM [PL_OPENSTOCK] WHERE [PL_KEY] = " & cPNL_Key & " AND [PL_SOURCENO] = '" & Trim(arrSource(J)) & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then arrPNL(J, 1) = CDbl(dr.Item(0))
            End If
            dr.Close()
            ''*** Purchase
            dr = datHandler.GetDataReader("SELECT SUM(cast([PL_AMOUNT] as money)) FROM [PL_PURCHASE] WHERE [PL_KEY] = " & cPNL_Key & " AND [PL_SOURCENO] = '" & Trim(arrSource(J)) & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then arrPNL(J, 2) = CDbl(dr.Item(0))
            End If
            dr.Close()
            ''*** Cost of Production
            dr = datHandler.GetDataReader("SELECT SUM(cast([EXA_AMOUNT] as money)) FROM [PL_PRODUCTION_COST] WHERE [EXA_KEY] = " & cPNL_Key & " AND [EXA_SOURCENO] = '" & Trim(arrSource(J)) & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then arrPNL(J, 3) = CDbl(dr.Item(0))
            End If
            dr.Close()
            ''*** Closing Stock
            dr = datHandler.GetDataReader("SELECT sum([PL_AMOUNT] as money)) FROM [PL_CLOSESTOCK] WHERE [PL_KEY] = " & cPNL_Key & " AND [PL_SOURCENO] = '" & Trim(arrSource(J)) & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then arrPNL(J, 4) = CDbl(dr.Item(0))
            End If
            dr.Close()

            ''Cost of sales (Opening Stock + Purchase + Cost of Production - Closing Stock)
            ''Gross Profit and Loss (Sales - Cost of Sales)
            arrPNL(J, 5) = arrPNL(J, 0) - ((arrPNL(J, 1)) + (arrPNL(J, 2)) + (arrPNL(J, 3)) - (arrPNL(J, 4)))

            If arrPNL(J, 5) > 0 Then
                osTotal = osTotal + arrPNL(J, 5)
            End If

        Next

eSub:
        OtherSource_GrossProfitLoss = osTotal

    End Function

    'NGOHCS B2011.2
    Protected Function GetStatusOfTax() As String
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim prmOledb(0) As SqlParameter
        Dim strStatus As String = ""

        ReDim prmOledb(1)
        prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
        prmOledb(1) = New SqlParameter("@ya", pdfForm.GetYA)
        ds = datHandler.GetData("SELECT TC_BALANCE_TAX_PAYABLE, TC_BALANCE_TAX_OVERPAID, TC_TAX_REPAYMENT FROM TAX_COMPUTATION WHERE TC_REF_NO=@ref_no AND TC_YA=@ya", prmOledb)
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
    Protected Function FormatDateSlash(ByVal dtTemp As String) As String

        Try
            Dim strTemp As String = ""

            If Not IsDBNull(dtTemp) Then
                strTemp = Format(CDate(dtTemp), "dd/MM/yyyy").ToString
            End If
            Return strTemp

        Catch ex As Exception
            Return "---"
        End Try

    End Function
    'Protected Function FormatDateSlash(ByVal dtTemp As Date) As String

    '    Dim strTemp As String = ""

    '    If Not IsDBNull(dtTemp) Then
    '        strTemp = Format(dtTemp, "dd/MM/yyyy").ToString
    '    End If
    '    Return strTemp

    'End Function

    Public Function FormatDateDash(ByVal strTemp As String) As String

        If Not Trim(strTemp) = "" Then
            strTemp = strTemp.Insert(2, "/").Insert(5, "/")
        End If
        Return strTemp

    End Function


    Protected Function FormatICNumber(ByVal strTemp As String) As String

        If Not Trim(strTemp) = "" Then
            strTemp = strTemp.Insert(6, "-").Insert(9, "-")
        End If
        Return strTemp

    End Function



#End Region
End Class

