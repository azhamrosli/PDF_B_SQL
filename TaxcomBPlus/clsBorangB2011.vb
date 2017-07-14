Imports iTextSharp.text.pdf
Imports System.Data.OleDb
Imports System.IO
Imports System

Public Class clsBorangB2011
    Private Const pdfSubFormName = "topmostSubform[0]."

    Dim pdfForm As New clsPDFMaker
    Dim pdfFormFields As AcroFields
    Dim datHandler As New clsDataHandler("")

#Region "CStor"

    Public Sub New()

        datHandler = New clsDataHandler(pdfForm.GetFormType)
        pdfFormFields = pdfForm.GetStamper.AcroFields
        CheckFieldEmpty()
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
        Page12()
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
        Dim strArray(1) As String

        ' ==== Master Data ==== "
        Try
            prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page3[0]."

            ds = datHandler.GetData("select tp_name , tp_ref_no_prefix, (tp_ref_no1 + tp_ref_no2 + tp_ref_no3)," _
                        & " (tp_ic_new_1 + tp_ic_new_2 + tp_ic_new_3), tp_ic_old," _
                        & " tp_police_no, tp_army_no, tp_passport_no," _
                        & " tp_country, tp_gender, tp_status, tp_date_marriage," _
                        & " tp_date_divorce, tp_type_assessment, tp_kup," _
                        & " (tp_curr_add_line1 + iif(tp_curr_add_line2 = '','',', ') + tp_curr_add_line2 + iif(tp_curr_add_line3 = '','',', ') + tp_curr_add_line3)," _
                        & " tp_curr_postcode, tp_curr_city, tp_curr_state," _
                        & " (tp_com_add_line1 + iif(tp_com_add_line2 = '','',', ') + tp_com_add_line2 + iif(tp_com_add_line3 = '','',', ') + tp_com_add_line3)," _
                        & " tp_com_postcode, tp_com_city, tp_com_state," _
                        & " tp_tel1, tp_tel2, tp_mobile1, tp_mobile2," _
                        & " (tp_employer_no2 + tp_employer_no3)," _
                        & " tp_email, tp_bank, tp_bank_acc, tp_assessmenton" _
                        & " from taxp_profile where (tp_ref_no1 + tp_ref_no2 + tp_ref_no3)=?", prmOledb)

            'weihong TP_WORKER_APPROVEDATE
            dr = datHandler.GetDataReader("select tp_last_passport_no, tp_bwa, TP_WORKER_APPROVEDATE, TP_COM_ADD_STATUS from taxp_profile2 where" _
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
                    If Not IsDBNull(dr("tp_last_passport_no")) Then
                        If Not String.IsNullOrEmpty(dr("tp_last_passport_no")) Then
                            pdfFormFields.SetField(pdfFieldPath & "VIII[0]", dr("tp_last_passport_no").ToString)
                        End If
                    End If
                    If Not IsDBNull(dr("tp_bwa")) Then
                        If Not String.IsNullOrEmpty(dr("tp_bwa")) Then
                            pdfFormFields.SetField(pdfFieldPath & "A13[0]", dr("tp_bwa").ToString)
                        End If
                    End If
                    'weihong
                    If Not IsDBNull(dr("TP_WORKER_APPROVEDATE")) Then
                        pdfFormFields.SetField(pdfFieldPath & "A9[0]", "1")
                        pdfFormFields.SetField(pdfFieldPath & "A9a[0]", FormatDate(dr("TP_WORKER_APPROVEDATE")))
                        ' pdfFormFields.SetField(pdfFieldPath & "A4[0]", FormatDate(ds.Tables(0).Rows(0).Item(11)))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A9[0]", "2")
                    End If
                    'weihong
                End If
                dr.Close()

                ' ==== PART A ==== "
                ReDim strArray(2)
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(8)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(8)) Then
                        pdfFormFields.SetField(pdfFieldPath & "A1[0]", ds.Tables(0).Rows(0).Item(8).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(9)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(9)) Then
                        pdfFormFields.SetField(pdfFieldPath & "A2[0]", ds.Tables(0).Rows(0).Item(9).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(10)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(10)) Then
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
                            If ds.Tables(0).Rows(0).Item(31).ToString = "1" Then
                                pdfFormFields.SetField(pdfFieldPath & "A5[0]", "1")
                            ElseIf ds.Tables(0).Rows(0).Item(31).ToString = "2" Then
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
                If Not String.IsNullOrEmpty(pdfForm.GetRecordKeep) Then
                    If pdfForm.GetRecordKeep.ToString = 1 Then
                        pdfFormFields.SetField(pdfFieldPath & "A7[0]", "1")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A7[0]", "2")
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(15)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(15).ToString()) Then
                        strArray = SplitText(ds.Tables(0).Rows(0).Item(15).ToString().Replace(",,", ",").Replace(", ,", ","), 26)
                        If Not String.IsNullOrEmpty(strArray(0)) Then
                            pdfFormFields.SetField(pdfFieldPath & "A8_1[0]", strArray(0).ToUpper)
                            pdfFormFields.SetField(pdfFieldPath & "A8_2[0]", strArray(1).ToUpper)
                            pdfFormFields.SetField(pdfFieldPath & "A8_3[0]", strArray(2).ToUpper)
                        End If
                    End If
                End If
                If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(16).ToString) Then
                    pdfFormFields.SetField(pdfFieldPath & "A8_4[0]", ds.Tables(0).Rows(0).Item(16).ToString.ToUpper)
                End If
                If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(17).ToString) Then
                    pdfFormFields.SetField(pdfFieldPath & "A8_5[0]", ds.Tables(0).Rows(0).Item(17).ToString.ToUpper)
                End If
                If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(18).ToString) Then
                    pdfFormFields.SetField(pdfFieldPath & "A8_6[0]", ds.Tables(0).Rows(0).Item(18).ToString.ToUpper)
                End If

                ReDim strArray(2)
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(19)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(19).ToString()) Then
                        strArray = SplitText(ds.Tables(0).Rows(0).Item(19).ToString().ToString().Replace(",,", ",").Replace(", ,", ","), 26)
                        If Not String.IsNullOrEmpty(strArray(0)) Then
                            pdfFormFields.SetField(pdfFieldPath & "A9_1[0]", strArray(0).ToUpper)
                            pdfFormFields.SetField(pdfFieldPath & "A9_2[0]", strArray(1).ToUpper)
                            pdfFormFields.SetField(pdfFieldPath & "A9_3[0]", strArray(2).ToUpper)
                        End If
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(20).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(20).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A9_4[0]", ds.Tables(0).Rows(0).Item(20).ToString.ToUpper)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(21).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(21).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A9_5[0]", ds.Tables(0).Rows(0).Item(21).ToString.ToUpper)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(22).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(22).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A9_6[0]", ds.Tables(0).Rows(0).Item(22).ToString.ToUpper)
                    End If
                End If

                pdfFormFields.SetField(pdfFieldPath & "A10[0]", FormatPhoneNumber( _
                                            ds.Tables(0).Rows(0).Item(23).ToString, _
                                            ds.Tables(0).Rows(0).Item(24).ToString, _
                                            ds.Tables(0).Rows(0).Item(25).ToString, _
                                            ds.Tables(0).Rows(0).Item(26).ToString))

                If Not IsDBNull(ds.Tables(0).Rows(0).Item(27).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(27).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A11[0]", ds.Tables(0).Rows(0).Item(27).ToString)
                    End If
                End If

                ReDim strArray(0)
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(28).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(28).ToString) Then
                        strArray = SplitText(ds.Tables(0).Rows(0).Item(28).ToString, 28)
                        pdfFormFields.SetField(pdfFieldPath & "A12[0]", strArray(0))
                    End If
                End If

                ReDim strArray(0)
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(29).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(29).ToString) Then
                        strArray = SplitText(ds.Tables(0).Rows(0).Item(29).ToString, 29)
                        pdfFormFields.SetField(pdfFieldPath & "A14[0]", strArray(0).ToUpper)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(30).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(30).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "A15[0]", ds.Tables(0).Rows(0).Item(30).ToString)
                    End If
                End If
            End If

            ''NGOHCS B2010.2
            ''Initialise
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
            ''NGOHCS B2010.2 END

            ReDim prmOledb(1)
            prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            prmOledb(1) = New OleDbParameter("@ya", pdfForm.GetYA)
            ds = datHandler.GetData("SELECT SUM(TCA_CBL) FROM TAX_ADJUSTED_LOSS WHERE TC_KEY IN (SELECT TC_KEY FROM TAX_COMPUTATION WHERE TC_REF_NO=? AND TC_YA=?)", prmOledb)

            If (ds.Tables(0).Rows.Count > 0) Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) Then
                    If (CDbl(ds.Tables(0).Rows(0).Item(0).ToString) > 0) Then
                        pdfFormFields.SetField(pdfFieldPath & "A7a[0]", "1")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A7a[0]", "2")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "A7a[0]", "2")
                End If
            Else
                pdfFormFields.SetField(pdfFieldPath & "A7a[0]", "2")
            End If
            ''NGOHCS B2010.2 END

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
        Dim strArray(1) As String
        Dim strHWIC As String = ""
        Dim intCounter As Integer = 1
        Dim intNumberRecord As Integer = 0
        Dim dblTotalIncome As Double = 0
        Dim dblRentalIncome As Double = 0

        Try
            prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page4[0]."


            ' ==== PART B ==== "
            ds = datHandler.GetData("select tp_hw_name, tp_hw_ref_no_prefix, tp_hw_ref_no1 , tp_hw_ref_no2 , tp_hw_ref_no3," _
                        & " (tp_hw_ic_new1 + tp_hw_ic_new2 + tp_hw_ic_new3), tp_hw_ic_old," _
                        & " tp_hw_police_no, tp_hw_army_no, tp_hw_passport_no" _
                        & " from taxp_profile where (tp_ref_no1 + tp_ref_no2 + tp_ref_no3)=?", prmOledb)

            ReDim strArray(1)

            If ds.Tables(0).Rows.Count > 0 Then
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(0).ToString()) Then
                        strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString(), 28)
                        If Not String.IsNullOrEmpty(strArray(0)) Then
                            pdfFormFields.SetField(pdfFieldPath & "B1_1[0]", strArray(0).ToUpper)
                            pdfFormFields.SetField(pdfFieldPath & "B1_2[0]", strArray(1).ToUpper)
                        End If
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(1)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(0).ToString()) Then
                        pdfFormFields.SetField(pdfFieldPath & "B2_1[0]", ds.Tables(0).Rows(0).Item(1).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(2)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(2).ToString) Then
                        strHWIC = strHWIC & ds.Tables(0).Rows(0).Item(2).ToString
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(3)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(3).ToString) Then
                        strHWIC = strHWIC & ds.Tables(0).Rows(0).Item(3).ToString
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(4)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(4).ToString) Then
                        strHWIC = strHWIC & ds.Tables(0).Rows(0).Item(4).ToString
                    End If
                End If
                If Not String.IsNullOrEmpty(strHWIC) Then
                    pdfFormFields.SetField(pdfFieldPath & "B2_2[0]", strHWIC.ToString)
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(5)) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(5).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B3[0]", ds.Tables(0).Rows(0).Item(5).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(6).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(6).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B4[0]", ds.Tables(0).Rows(0).Item(6).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(7).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(7).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B5[0]", ds.Tables(0).Rows(0).Item(7).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(8).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(8).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B6[0]", ds.Tables(0).Rows(0).Item(8).ToString)
                    End If
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item(9).ToString) Then
                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item(9).ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "B7[0]", ds.Tables(0).Rows(0).Item(9).ToString)
                    End If
                End If
            End If
            ds.Clear()

            ' ==== PART C ==== "
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
                        pdfFormFields.SetField(pdfFieldPath & "C" & intCounter.ToString & "_2[0]", FormatFixedAmount(dr("adjsi_net_stat_income").ToString))
                        dblTotalIncome = dblTotalIncome + CDbl(dr("adjsi_net_stat_income"))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "C" & intCounter.ToString & "_2[0]", 0)
                    End If
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
                    pdfFormFields.SetField(pdfFieldPath & "C" & (intNumberRecord + 1).ToString & "_2[0]", FormatFixedAmount(dblRentalIncome.ToString))
                End If
            Else
                If intNumberRecord = 3 And dblRentalIncome = 0 Then
                    dr = datHandler.GetDataReader("select top 3 adjsi_net_stat_income , adj_business from income_adjusted where" _
                                & " adj_ref_no='" & pdfForm.GetRefNo & "' and adj_ya= '" & pdfForm.GetYA & "' order by adj_business desc")
                    If dr.Read() Then
                        dr2 = datHandler.GetDataReader("select bc_code from business_source where " _
                                    & " bc_key='" & pdfForm.GetRefNo & "' and bc_ya='" & pdfForm.GetYA & "'" _
                                    & " and bc_businesssource='" & dr("adj_business") & "'")
                        If dr2.Read() Then
                            If Not IsDBNull(dr2("bc_code")) Then
                                pdfFormFields.SetField(pdfFieldPath & "C3_1[0]", dr2("bc_code").ToString)
                            End If
                            If Not IsDBNull(dr("adjsi_net_stat_income")) Then
                                pdfFormFields.SetField(pdfFieldPath & "C3_2[0]", FormatFixedAmount(dr("adjsi_net_stat_income").ToString))
                            Else
                                pdfFormFields.SetField(pdfFieldPath & "C3_2[0]", 0)
                            End If
                        End If
                        dr2.Close()
                    End If
                    dr.Close()

                Else
                    dr = datHandler.GetDataReader("select sum(cdbl(adjsi_net_stat_income)) from income_adjusted where " _
                                   & "adj_ref_no ='" & pdfForm.GetRefNo & "' and adj_ya= '" & pdfForm.GetYA & "'")
                    If dr.Read() Then
                        If Not IsDBNull(dr.Item(0)) Then
                            dblTotalIncome = CDbl(dr.Item(0)) - dblTotalIncome
                        End If
                    End If
                    dblTotalIncome = dblTotalIncome + dblRentalIncome
                    If dblTotalIncome > 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "C3_2[0]", FormatFixedAmount(dblTotalIncome))
                    End If
                    dr.Close()
                End If
            End If

            'dr = Nothing
            'dr2 = Nothing
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
                        pdfFormFields.SetField(pdfFieldPath & "C" & intCounter.ToString & "_1[0]", FormatFixedAmount(dr2.Item(0)))
                    End If
                End If
                If Not IsDBNull(dr("ps_sch_7a_stat_income")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C" & intCounter.ToString & "_2[0]", FormatFixedAmount(dr("ps_sch_7a_stat_income")))
                    dblTotalIncome = dblTotalIncome + CDbl(dr("ps_sch_7a_stat_income"))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "C" & intCounter.ToString & "_2[0]", 0)
                End If
                intCounter = intCounter + 1
                dr2.Close()
            End While
            'Else
            'End If
            dr.Close()

            If intNumberRecord > 3 Then
                dr = datHandler.GetDataReader("select sum(cdbl(ps_sch_7a_stat_income)) from income_partnership where " _
                                   & " pn_ref_no='" & pdfForm.GetRefNo & "'and pn_ya= '" & pdfForm.GetYA & "'")

                If dr.Read() Then

                    If IsDBNull(dr.Item(0)) Or String.IsNullOrEmpty(dr.Item(0)) Then
                        dblTotalIncome = 0 - dblTotalIncome
                    Else
                        dblTotalIncome = dr.Item(0) - dblTotalIncome
                    End If
                    If dblTotalIncome >= 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "C6_2[0]", FormatFixedAmount(dblTotalIncome))
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
                                & " TC_TOTAL1 from tax_computation where" _
                                & " tc_ref_no ='" & pdfForm.GetRefNo & "' and tc_ya ='" & pdfForm.GetYA & "'")

            If dr.Read() Then
                If Not IsDBNull(dr("TC_STATUTORY_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C7[0]", FormatFixedAmount(dr("TC_STATUTORY_INCOME")))
                End If
                If Not IsDBNull(dr("TC_BUSINESSLOSS_BF")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C8[0]", FormatFixedAmount(dr("TC_BUSINESSLOSS_BF")))
                End If
                If Not IsDBNull(dr("TC_AGGREGATE_BUS_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C9[0]", FormatFixedAmount(dr("TC_AGGREGATE_BUS_INCOME")))
                End If
                If Not IsDBNull(dr("TC_EMPLOYMENT_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C10[0]", FormatFixedAmount(dr("TC_EMPLOYMENT_INCOME")))
                End If
                If Not IsDBNull(dr("TC_DIVIDEND")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C11[0]", FormatFixedAmount(dr("TC_DIVIDEND")))
                End If
                If Not IsDBNull(dr.Item(5)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C12[0]", FormatFixedAmount(dr.Item(5).ToString))
                End If
                If Not IsDBNull(dr.Item(6)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C13[0]", FormatFixedAmount(dr.Item(6).ToString))
                End If
                If Not IsDBNull(dr("TC_PENSION_AND_ETC")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C14[0]", FormatFixedAmount(dr("TC_PENSION_AND_ETC")))
                End If
                If Not IsDBNull(dr.Item(8)) Then
                    pdfFormFields.SetField(pdfFieldPath & "C15[0]", FormatFixedAmount(dr.Item(8).ToString))
                End If
                If Not IsDBNull(dr("TC_ADDITION_43")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C16[0]", FormatFixedAmount(dr("TC_ADDITION_43")))
                End If
                If Not IsDBNull(dr("TC_AGGREGATE_OTHER_SRC")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C17[0]", FormatFixedAmount(dr("TC_AGGREGATE_OTHER_SRC")))
                End If
                If Not IsDBNull(dr("TC_AGGREGATE_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C18[0]", FormatFixedAmount(dr("TC_AGGREGATE_INCOME")))
                End If
                If Not IsDBNull(dr("TC_BUSINESSLOSS_CY")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C19[0]", FormatFixedAmount(dr("TC_BUSINESSLOSS_CY")))
                End If
                If Not IsDBNull(dr("TC_TOTAL1")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C20[0]", FormatFixedAmount(dr("TC_TOTAL1")))
                End If
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
        Dim dblTotalValue1 As Double = 0
        Dim intCounter As Integer = 1

        Try
            ' prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page5[0]."


            ' ==== PART C ==== "
            'C21 - C23
            dr = datHandler.GetDataReader("select TC_KEY, TC_AGGREGATE_INCOME, TC_PROSPECTING, TC_QUALIFYING_AG_EXP, TC_TOTAL2," _
                                    & " TC_4, TC_3, TC_TOTAL_INCOME_2, TC_INCOME_TRANSFER_FROM_HW, TC_TOTAL_INCOME_3" _
                                    & " from tax_computation where" _
                                    & " tc_ref_no ='" & pdfForm.GetRefNo & "' and tc_ya ='" & pdfForm.GetYA & "'")
            If dr.Read Then

                If Not IsDBNull(dr("TC_PROSPECTING")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C21[0]", FormatFixedAmount(dr("TC_PROSPECTING")))
                End If
                If Not IsDBNull(dr("TC_QUALIFYING_AG_EXP")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C22[0]", FormatFixedAmount(dr("TC_QUALIFYING_AG_EXP")))
                End If
                If Not IsDBNull(dr("TC_TOTAL2")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C23[0]", FormatFixedAmount(dr("TC_TOTAL2")))
                End If

                'C24 - 'C31
                dr2 = datHandler.GetDataReader("select TCG_KEY, TCG_AMOUNT" _
                                        & " from tax_gifts where" _
                                        & " tc_key =" & dr("TC_KEY"))
                Do While dr2.Read()
                    If Not IsDBNull(dr2("TCG_KEY")) Then

                        Select Case dr2("TCG_KEY")
                            Case "9"
                                pdfFormFields.SetField(pdfFieldPath & "C24[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                            Case "1"
                                pdfFormFields.SetField(pdfFieldPath & "C24A[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "7"
                                pdfFormFields.SetField(pdfFieldPath & "C25[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "8"
                                pdfFormFields.SetField(pdfFieldPath & "C26_1[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                                dblTotalValue = dblTotalValue + dr2("TCG_AMOUNT")
                            Case "2"
                                pdfFormFields.SetField(pdfFieldPath & "C27[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                            Case "3"
                                pdfFormFields.SetField(pdfFieldPath & "C28[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                            Case "4"
                                pdfFormFields.SetField(pdfFieldPath & "C29[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                            Case "5"
                                pdfFormFields.SetField(pdfFieldPath & "C30[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))
                            Case "6"
                                pdfFormFields.SetField(pdfFieldPath & "C31[0]", FormatFixedAmount(dr2("TCG_AMOUNT")))

                        End Select
                    End If
                Loop
                dr2.Close()

                'C26 Total restrict to 7% of C28
                If Not IsDBNull(dr("TC_AGGREGATE_INCOME")) Then
                    If Not String.IsNullOrEmpty(dr("TC_AGGREGATE_INCOME").ToString) Then
                        If dblTotalValue >= CDbl(dr("TC_AGGREGATE_INCOME")) * 0.07 Then
                            pdfFormFields.SetField(pdfFieldPath & "C26_2[0]", FormatFixedAmount((CDbl(dr("TC_AGGREGATE_INCOME")) * 0.07).ToString))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "C26_2[0]", FormatFixedAmount(dblTotalValue.ToString))
                        End If
                    End If
                End If

                'C32 - C35_2

                If Not IsDBNull(dr("TC_4")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C32[0]", FormatFixedAmount(dr("TC_4")))
                End If

                If Not IsDBNull(dr("TC_3")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C33[0]", FormatFixedAmount(dr("TC_3")))
                End If

                If Not IsDBNull(dr("TC_3")) And Not IsDBNull(dr("TC_4")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C34[0]", FormatFloatingAmount(CDbl(dr("TC_3")) + CDbl(dr("TC_4")), False))
                ElseIf Not IsDBNull(dr("TC_4")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C34[0]", dr("TC_4"))
                ElseIf Not IsDBNull(dr("TC_3")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C34[0]", dr("TC_3"))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "C34[0]", "0")
                End If


                If Not IsDBNull(dr("TC_INCOME_TRANSFER_FROM_HW")) Then
                    pdfFormFields.SetField(pdfFieldPath & "C35_2[0]", FormatFixedAmount(dr("TC_INCOME_TRANSFER_FROM_HW")))
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
                    'If Not IsDBNull(dr("TC_3")) And Not IsDBNull(dr("TC_4")) Then
                    '    pdfFormFields.SetField(pdfFieldPath & "C36[0]", (CDbl(dr("TC_3")) + CDbl(dr("TC_4")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))))
                    'ElseIf Not IsDBNull(dr("TC_4")) Then
                    '    pdfFormFields.SetField(pdfFieldPath & "C36[0]", (CDbl(dr("TC_4")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))))
                    'ElseIf Not IsDBNull(dr("TC_3")) Then
                    '    pdfFormFields.SetField(pdfFieldPath & "C36[0]", (CDbl(dr("TC_3")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))))
                    'Else
                    '    pdfFormFields.SetField(pdfFieldPath & "C36[0]", CDbl(dr("TC_INCOME_TRANSFER_FROM_HW")))
                    'End If

                    If Not IsDBNull(dr2("TP_STATUS")) Then
                        If dr2("TP_STATUS") = "1" Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        ElseIf dr2("TP_STATUS") = "2" And CDbl(pdfForm.GetYA) >= 2007 And dr2("TP_TYPE_ASSESSMENT") = "3" Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        ElseIf dr2("TP_STATUS") = "2" And CDbl(pdfForm.GetYA) = 2006 And dr2("TP_TYPE_ASSESSMENT") = "1" And CDbl(dr("TC_INCOME_TRANSFER_FROM_HW")) = 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
                        ElseIf (dr2("TP_GENDER") = "1" And dr2("TP_STATUS") = "2" And dr2("TP_TYPE_ASSESSMENT") = "1" And _
                            dr2("TP_ASSESSMENTON") = "1") Or (dr2("TP_GENDER") = "2" And dr2("TP_STATUS") = "2" And dr2("TP_TYPE_ASSESSMENT") = "1" And dr2("TP_ASSESSMENTON") = "2") Then
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", FormatFixedAmount((CDbl(dr("TC_TOTAL_INCOME_2")) + CDbl(dr("TC_INCOME_TRANSFER_FROM_HW"))).ToString))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "C36[0]", "0")
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
                    If Not IsDBNull(dr2("tcc_amount")) Then
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
                                    'weihong
                                    pdfFormFields.SetField(pdfFieldPath & "D12[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
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
                                Case 9
                                    pdfFormFields.SetField(pdfFieldPath & "D8D[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                Case 25
                                    pdfFormFields.SetField(pdfFieldPath & "D18[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                                    dblTotalValue1 = dblTotalValue1 + CDbl(dr2("tcc_amount"))
                            End Select
                        End If
                    End If
                End While
                'weihong
                pdfFormFields.SetField(pdfFieldPath & "D7_2[0]", FormatFixedAmount(dblTotalValue.ToString))
                pdfFormFields.SetField(pdfFieldPath & "D17[0]", FormatFixedAmount(dblTotalValue1.ToString))
                'pdfFormFields.SetField(pdfFieldPath & "D12[0]", FormatFixedAmount(dr2("tcc_amount").ToString))
                dr2.Close()
            End If
            dr.Close()


            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama5", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo.ToString) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj5", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Page4()
        'weihong
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
            pdfFieldPath = pdfSubFormName & "Page6[0]."

            ' ==== Part E ==== '
            'E1 - E15 exclude E4 - E7
            dr = datHandler.GetDataReader("Select TC_KEY, TC_RELIEF, TC_CHARGEABLE_INCOME, TC_TAX_FIRST_INCOME, TC_TAX_FIRST_TAX," _
                                & " TC_TAX_BALANCE_INCOME, TC_TAX_BALANCE_RATE, TC_TAX_BALANCE_TAX, TC_TOTAL_INCOME_TAX," _
                                & " TC_REBATES, TC_INCOME_TAX_CHARGED, TC_SEC110_DIVIDEND,TC_SEC110_OTHERS, TC_1," _
                                & " TC_2, TC_TAX_PAYABLE, TC_TAX_REPAYMENT, TC_TAX_SCH1_INCOME, TC_TAX_SCH1_TAX" _
                                & " From TAX_COMPUTATION Where" _
                                & " TC_REF_NO= '" & pdfForm.GetRefNo & "' and TC_YA= '" & pdfForm.GetYA & "'")

            '& " TC_REF_NO= '" & pdfForm.GetRefNo & "' and TC_YA= '" & pdfForm.GetYA & "'")
            '& " EI_REF_NO= '" & pdfForm.GetRefNo & "' and EI_YA= '" & pdfForm.GetYA & "'")

            If dr.Read() Then
                'D14
                If Not IsDBNull(dr("TC_TAX_SCH1_INCOME")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E2_1[0]", FormatFixedAmount(dr("TC_TAX_SCH1_INCOME").ToString))
                End If
                If Not IsDBNull(dr("TC_TAX_SCH1_TAX")) Then
                    pdfFormFields.SetField(pdfFieldPath & "E2_2[0]", FormatFloatingAmount(dr("TC_TAX_SCH1_TAX").ToString, True))
                End If
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
                        pdfFormFields.SetField(pdfFieldPath & "Nama6", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj6", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub Page5()

        Dim pdfFieldPath As String
        Dim dr As OleDbDataReader = Nothing
        Dim dr2 As OleDbDataReader = Nothing
        Dim dblTemp As Double = 0
        Dim dblTotalValue As Double = 0
        Dim dblTotalValue2 As Double = 0
        Dim intCounter As Integer = 1
        Dim strArray(1) As String


        Try
            pdfFieldPath = pdfSubFormName & "Page7[0]."


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
            dr = datHandler.GetDataReader("Select * from preceding_year where py_ref_no= '" & pdfForm.GetRefNo & "' and py_ya= '" & pdfForm.GetYA & "'")
            'initialise with 0 for Amaun Kasar, Caruman Kumpulan Wang
            pdfFormFields.SetField(pdfFieldPath & "G1_3[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "G1_4[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "G2_3[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "G2_4[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "G3_3[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "G3_4[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "G4_3[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "G4_4[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "G5_3[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "G5_4[0]", 0)
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
                    End If
                    If Not IsDBNull(dr2("PY_EPF")) Then
                        pdfFormFields.SetField(pdfFieldPath & "G" & intCounter.ToString & "_4[0]", FormatFixedAmount(dr2("PY_EPF").ToString))
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
                    If Not String.IsNullOrEmpty(dr("TP_ADM_NAME").ToString) Then
                        strArray = SplitText(dr("TP_ADM_NAME").ToString, 28)
                        If Not String.IsNullOrEmpty(strArray(0)) Then
                            pdfFormFields.SetField(pdfFieldPath & "H1_1[0]", strArray(0).ToString.ToUpper)
                            pdfFormFields.SetField(pdfFieldPath & "H1_2[0]", strArray(1).ToString.ToUpper)
                        End If
                    End If
                End If
                If Not String.IsNullOrEmpty(dr.Item(1)) Then 'weihong
                    pdfFormFields.SetField(pdfFieldPath & "H2[0]", dr.Item(1).ToString)
                End If
                If Not String.IsNullOrEmpty(dr("TP_ADM_IC_OLD")) Then 'weihong
                    pdfFormFields.SetField(pdfFieldPath & "H3[0]", dr("TP_ADM_IC_OLD").ToString)
                End If
                If Not String.IsNullOrEmpty(dr("TP_ADM_POLICE_NO")) Then 'weihong
                    pdfFormFields.SetField(pdfFieldPath & "H4[0]", dr("TP_ADM_POLICE_NO").ToString)
                End If
                If Not String.IsNullOrEmpty(dr("TP_ADM_ARMY_NO")) Then 'weihong
                    pdfFormFields.SetField(pdfFieldPath & "H5[0]", dr("TP_ADM_ARMY_NO").ToString)
                End If
                If Not String.IsNullOrEmpty(dr("TP_ADM_PASSPORT_NO")) Then 'weihong
                    pdfFormFields.SetField(pdfFieldPath & "H6[0]", dr("TP_ADM_PASSPORT_NO").ToString)
                End If
            End If
            dr.Close()

            ' === Part J === '
            'J1 - J1d
            dr = datHandler.GetDataReader("Select TC_AL_CY_UNASORBED_LOSS, TC_AL_BAL_UNASORBED_LOSS, TC_AL_BALANCE_CF," _
                                    & " TC_PIONEER, TC_PIONEER_CF" _
                                    & " From TAX_COMPUTATION Where" _
                                    & " TC_REF_NO='" & pdfForm.GetRefNo & "' AND TC_YA='" & pdfForm.GetYA & "'")
            If dr.Read() Then
                If Not IsDBNull(dr("TC_AL_CY_UNASORBED_LOSS")) Then
                    pdfFormFields.SetField(pdfFieldPath & "J1a2[0]", FormatFixedAmount(dr("TC_AL_CY_UNASORBED_LOSS").ToString))
                End If
                If Not IsDBNull(dr("TC_AL_BAL_UNASORBED_LOSS")) Then
                    pdfFormFields.SetField(pdfFieldPath & "J1b[0]", FormatFixedAmount(dr("TC_AL_BAL_UNASORBED_LOSS").ToString))
                End If
                If Not IsDBNull(dr("TC_AL_BALANCE_CF")) Then
                    pdfFormFields.SetField(pdfFieldPath & "J1c[0]", FormatFixedAmount(dr("TC_AL_BALANCE_CF").ToString))
                End If
                If Not IsDBNull(dr("TC_PIONEER")) Then
                    pdfFormFields.SetField(pdfFieldPath & "J1d_1[0]", FormatFixedAmount(dr("TC_PIONEER").ToString))
                End If
                If Not IsDBNull(dr("TC_PIONEER_CF")) Then
                    pdfFormFields.SetField(pdfFieldPath & "J1d_2[0]", FormatFixedAmount(dr("TC_PIONEER_CF").ToString))
                End If
            End If
            dr.Close()

            dr = datHandler.GetDataReader("select sum(tca_cbl) from tax_adjusted_loss where tc_key in " _
                        & "(select tc_key from tax_computation where " _
                        & "tc_ref_no ='" & pdfForm.GetRefNo & "' AND TC_YA='" & pdfForm.GetYA & "')")
            If dr.Read() Then
                If Not IsDBNull(dr(0)) Then
                    pdfFormFields.SetField(pdfFieldPath & "J1a1[0]", FormatFixedAmount(dr(0).ToString))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "J1a1[0]", "0")
                End If
            Else
                pdfFormFields.SetField(pdfFieldPath & "J1a1[0]", "0")
            End If
            dr.Close()

            'J2a - J2b
            intCounter = 1
            dr = datHandler.GetDataReader("SELECT TOP 2 ADCA_UTIL ,ADCA_BAL_CF" _
                                    & " From INCOME_ADJUSTED WHERE" _
                                    & " ADJ_REF_NO= '" & pdfForm.GetRefNo & "' AND ADJ_YA = '" & pdfForm.GetYA & "'" _
                                    & " ORDER BY ADJ_BUSINESS")
            pdfFormFields.SetField(pdfFieldPath & "J2a_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J2a_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J2b_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J2b_2[0]", 0)
            'If dr.Read Then
            While dr.Read()
                If Not IsDBNull(dr("ADCA_UTIL") And Not IsDBNull("ADCA_BAL_CF")) Then
                    dblTotalValue = dblTotalValue + CDbl(dr("ADCA_UTIL"))
                    dblTotalValue2 = dblTotalValue2 + CDbl(dr("ADCA_BAL_CF"))
                    Select Case intCounter
                        Case 1
                            pdfFormFields.SetField(pdfFieldPath & "J2a_1[0]", FormatFixedAmount(dr("ADCA_UTIL").ToString))
                            pdfFormFields.SetField(pdfFieldPath & "J2a_2[0]", FormatFixedAmount(dr("ADCA_BAL_CF").ToString))
                        Case 2
                            pdfFormFields.SetField(pdfFieldPath & "J2b_1[0]", FormatFixedAmount(dr("ADCA_UTIL").ToString))
                            pdfFormFields.SetField(pdfFieldPath & "J2b_2[0]", FormatFixedAmount(dr("ADCA_BAL_CF").ToString))
                    End Select
                Else
                    Select Case intCounter
                        Case 1
                            pdfFormFields.SetField(pdfFieldPath & "J2a_1[0]", 0)
                            pdfFormFields.SetField(pdfFieldPath & "J2a_2[0]", 0)
                        Case 2
                            pdfFormFields.SetField(pdfFieldPath & "J2b_1[0]", 0)
                            pdfFormFields.SetField(pdfFieldPath & "J2b_2[0]", 0)
                    End Select
                End If
                intCounter = intCounter + 1
            End While
            'Else
            'End If
            dr.Close()


            'J2c
            dr = datHandler.GetDataReader("SELECT SUM(CDBL(ADCA_UTIL)),SUM(CDBL(ADCA_BAL_CF))" _
                                    & " From INCOME_ADJUSTED WHERE" _
                                    & " ADJ_REF_NO= '" & pdfForm.GetRefNo & "' AND ADJ_YA = '" & pdfForm.GetYA & "'")

            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) And Not IsDBNull(dr.Item(1)) Then
                    dblTotalValue = CDbl(dr.Item(0)) - dblTotalValue
                    dblTotalValue2 = CDbl(dr.Item(1)) - dblTotalValue2
                    pdfFormFields.SetField(pdfFieldPath & "J2c_1[0]", FormatFixedAmount(dblTotalValue.ToString))
                    pdfFormFields.SetField(pdfFieldPath & "J2c_2[0]", FormatFixedAmount(dblTotalValue2.ToString))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "J2c_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "J2c_2[0]", 0)
                End If
            Else
                pdfFormFields.SetField(pdfFieldPath & "J2c_1[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "J2c_2[0]", 0)
            End If
            dr.Close()

            'J2d - J2e
            intCounter = 1
            dblTotalValue = 0
            dblTotalValue2 = 0
            dr = datHandler.GetDataReader("SELECT TOP 2 PSCA_UTIL ,PSCA_BAL_CF" _
                                    & " From INCOME_PARTNERSHIP WHERE" _
                                    & " PN_REF_NO= '" & pdfForm.GetRefNo & "' AND PN_YA = '" & pdfForm.GetYA & "'" _
                                    & " ORDER BY PS_SOURCE")
            pdfFormFields.SetField(pdfFieldPath & "J2d_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J2d_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J2e_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J2e_2[0]", 0)
            'If dr.Read Then
            While dr.Read()
                If Not IsDBNull(dr("PSCA_UTIL") And Not IsDBNull("PSCA_BAL_CF")) Then
                    dblTotalValue = dblTotalValue + CDbl(dr("PSCA_UTIL"))
                    dblTotalValue2 = dblTotalValue2 + CDbl(dr("PSCA_BAL_CF"))
                    Select Case intCounter
                        Case 1
                            pdfFormFields.SetField(pdfFieldPath & "J2d_1[0]", FormatFixedAmount(dr("PSCA_UTIL").ToString))
                            pdfFormFields.SetField(pdfFieldPath & "J2d_2[0]", FormatFixedAmount(CDbl(dr("PSCA_BAL_CF")).ToString))
                        Case 2
                            pdfFormFields.SetField(pdfFieldPath & "J2e_1[0]", FormatFixedAmount(dr("PSCA_UTIL").ToString))
                            pdfFormFields.SetField(pdfFieldPath & "J2e_2[0]", FormatFixedAmount(CDbl(dr("PSCA_BAL_CF")).ToString))
                    End Select
                Else
                    Select Case intCounter
                        Case 1
                            pdfFormFields.SetField(pdfFieldPath & "J2d_1[0]", 0)
                            pdfFormFields.SetField(pdfFieldPath & "J2d_2[0]", 0)
                        Case 2
                            pdfFormFields.SetField(pdfFieldPath & "J2e_1[0]", 0)
                            pdfFormFields.SetField(pdfFieldPath & "J2e_2[0]", 0)
                    End Select
                End If
                intCounter = intCounter + 1
            End While
            'Else
            'End If
            dr.Close()


            'J2f
            dr = datHandler.GetDataReader("SELECT SUM(CDBL(PSCA_UTIL)),SUM(CDBL(PSCA_BAL_CF))" _
                                    & " From INCOME_PARTNERSHIP WHERE" _
                                    & " PN_REF_NO= '" & pdfForm.GetRefNo & "' AND PN_YA = '" & pdfForm.GetYA & "'")

            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) And Not IsDBNull(dr.Item(1)) Then
                    dblTotalValue = CDbl(dr.Item(0)) - dblTotalValue
                    dblTotalValue2 = CDbl(dr.Item(1)) - dblTotalValue2
                    If dblTotalValue >= 0 Or dblTotalValue2 >= 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "J2f_1[0]", FormatFixedAmount(dblTotalValue.ToString))
                        pdfFormFields.SetField(pdfFieldPath & "J2f_2[0]", FormatFixedAmount(dblTotalValue2.ToString))
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "J2f_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "J2f_2[0]", 0)
                End If
            Else
                pdfFormFields.SetField(pdfFieldPath & "J2f_1[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "J2f_2[0]", 0)
            End If
            dr.Close()


            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama7", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj7", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try


    End Sub

    Private Sub Page6()

        Dim pdfFieldPath As String
        Dim dr As OleDbDataReader = Nothing
        Dim dr2 As OleDbDataReader = Nothing
        Dim dblTemp As Double = 0
        Dim dblTotalValue As Double = 0
        Dim dblTotalValue2 As Double = 0
        Dim dblTotalValue3 As Double = 0
        Dim intCounter As Integer = 1
        Dim strArray(1) As String
        'NGOHCS B+ C2009.1 (SU11)
        Dim boolHasRecord As Boolean = False
        'NGOHCS B+ C2009.1 (SU11) END

        Try
            ' prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page8[0]."


            ' === Part J === '

            'J3a - J3d
            dr = datHandler.GetDataReader("SELECT NR_SECTION, NR_GROSS_TOTAL, NR_WITHHOLD, NR_WITHHOLD_107A" _
                                                & " From NON_RESIDENT WHERE" _
                                                & " NR_REF_NO= '" & pdfForm.GetRefNo & "' AND NR_YA = '" & pdfForm.GetYA & "'" _
                                                & " ORDER BY NR_SECTION")
            pdfFormFields.SetField(pdfFieldPath & "J3a_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J3a_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J3b_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J3b_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J3c_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J3c_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J3d_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J3d_2[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J3e_1[0]", 0)
            pdfFormFields.SetField(pdfFieldPath & "J3e_2[0]", 0)
            'If dr.Read Then
            While dr.Read()
                If Not IsDBNull(dr("NR_GROSS_TOTAL")) And Not IsDBNull("NR_WITHHOLD") And Not IsDBNull("NR_SECTION") Then
                    Select Case dr("NR_SECTION")
                        Case 1
                            pdfFormFields.SetField(pdfFieldPath & "J3a_1[0]", FormatFixedAmount(dr("NR_GROSS_TOTAL").ToString))
                            pdfFormFields.SetField(pdfFieldPath & "J3a_2[0]", FormatFixedAmount((CDbl(dr("NR_WITHHOLD")) + CDbl(dr("NR_WITHHOLD_107A"))).ToString))
                        Case 2
                            pdfFormFields.SetField(pdfFieldPath & "J3b_1[0]", FormatFixedAmount(dr("NR_GROSS_TOTAL").ToString))
                            pdfFormFields.SetField(pdfFieldPath & "J3b_2[0]", FormatFixedAmount(dr("NR_WITHHOLD").ToString))
                        Case 3
                            pdfFormFields.SetField(pdfFieldPath & "J3c_1[0]", FormatFixedAmount(dr("NR_GROSS_TOTAL").ToString))
                            pdfFormFields.SetField(pdfFieldPath & "J3c_2[0]", FormatFixedAmount(dr("NR_WITHHOLD").ToString))
                        Case 4
                            pdfFormFields.SetField(pdfFieldPath & "J3d_1[0]", FormatFixedAmount(dr("NR_GROSS_TOTAL").ToString))
                            pdfFormFields.SetField(pdfFieldPath & "J3d_2[0]", FormatFixedAmount(dr("NR_WITHHOLD").ToString))
                        Case 6
                            pdfFormFields.SetField(pdfFieldPath & "J3e_1[0]", FormatFixedAmount(dr("NR_GROSS_TOTAL").ToString))
                            pdfFormFields.SetField(pdfFieldPath & "J3e_2[0]", FormatFixedAmount(dr("NR_WITHHOLD").ToString))
                    End Select
                End If
            End While
            'Else
            'End If
            dr.Close()


            'K1 - K4
            intCounter = 1
            dblTotalValue = 0
            dr = datHandler.GetDataReader("Select ADJ_KEY" _
                                                & " From INCOME_ADJUSTED WHERE" _
                                                & " ADJ_REF_NO= '" & pdfForm.GetRefNo & "' AND ADJ_YA = '" & pdfForm.GetYA & "'")

            'NGOHCS B+ C2009.1 (SU11)
            'If dr.Read Then
            While dr.Read()
                dr2 = datHandler.GetDataReader("Select ADJD_CLAIM_CODE, ADJD_AMOUNT" _
                                                & " From INCOME_ADJ_FURTHER Where" _
                                                & " ADJ_KEY= " & dr("ADJ_KEY") & " Order By ADJD_ID, ADJD_NO")
                While dr2.Read()
                    boolHasRecord = True
                    If Not IsDBNull(dr2("ADJD_CLAIM_CODE")) And Not IsDBNull(dr2("ADJD_AMOUNT")) Then
                        If intCounter <= 4 Then
                            pdfFormFields.SetField(pdfFieldPath & "K" & intCounter.ToString & "_1[0]", dr2("ADJD_CLAIM_CODE").ToString)
                            pdfFormFields.SetField(pdfFieldPath & "K" & intCounter.ToString & "_2[0]", FormatFixedAmount(dr2("ADJD_AMOUNT").ToString))
                        End If
                        dblTotalValue = dblTotalValue + CDbl(dr2("ADJD_AMOUNT"))
                        intCounter = intCounter + 1
                    Else
                        If intCounter <= 4 Then
                            pdfFormFields.SetField(pdfFieldPath & "K" & intCounter.ToString & "_1[0]", 0)
                            pdfFormFields.SetField(pdfFieldPath & "K" & intCounter.ToString & "_2[0]", 0)
                        End If
                    End If
                End While
                dr2.Close()
            End While

            If intCounter <= 4 Then
                For i As Integer = intCounter To 4
                    'pdfFormFields.SetField(pdfFieldPath & "K" & intCounter.ToString & "_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "K" & i.ToString & "_2[0]", 0)
                Next
            End If

            If Not boolHasRecord Then
                pdfFormFields.SetField(pdfFieldPath & "K1_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "K2_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "K3_2[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "K4_2[0]", 0)
            End If
            'End If
            'NGOHCS B+ C2009.1 (SU11) END
            dr.Close()

            'K5
            pdfFormFields.SetField(pdfFieldPath & "K5[0]", FormatFixedAmount(dblTotalValue.ToString))

            ' === Part L === '

            'L1 - L2

            dr = datHandler.GetDataReader("Select TC_KEY From TAX_COMPUTATION Where" _
                            & " TC_REF_NO= '" & pdfForm.GetRefNo & "' And TC_YA= '" & pdfForm.GetYA & "'")

            If dr.Read() Then

                dr2 = datHandler.GetDataReader("Select TIC_KEY, TIC_CLAIM, TIC_CF" _
                            & " From TAX_INCENTIVE_CLAIM Where" _
                            & " TC_KEY= " & dr("TC_KEY") & " Order By TIC_KEY")

                While dr2.Read()
                    If Not IsDBNull(dr2("TIC_KEY")) Then
                        Select Case dr2("TIC_KEY")
                            Case 1
                                If Not IsDBNull(dr2("TIC_CLAIM")) Then
                                    If Not String.IsNullOrEmpty(dr2("TIC_CLAIM").ToString) Then
                                        pdfFormFields.SetField(pdfFieldPath & "L1_1[0]", FormatFixedAmount(dr2("TIC_CLAIM").ToString))
                                    End If
                                End If
                                If Not IsDBNull(dr2("TIC_CF")) Then
                                    If Not String.IsNullOrEmpty(dr2("TIC_CF").ToString) Then
                                        pdfFormFields.SetField(pdfFieldPath & "L1_2[0]", FormatFixedAmount(dr2("TIC_CF").ToString))
                                    End If
                                End If
                            Case 2
                                If Not IsDBNull(dr2("TIC_CLAIM")) Then
                                    If Not String.IsNullOrEmpty(dr2("TIC_CLAIM").ToString) Then
                                        pdfFormFields.SetField(pdfFieldPath & "L2_1[0]", FormatFixedAmount(dr2("TIC_CLAIM").ToString))
                                    End If
                                End If
                                If Not IsDBNull(dr2("TIC_CF")) Then
                                    If Not String.IsNullOrEmpty(dr2("TIC_CF").ToString) Then
                                        pdfFormFields.SetField(pdfFieldPath & "L2_2[0]", FormatFixedAmount(dr2("TIC_CF").ToString))
                                    End If
                                End If
                            Case 3
                                If Not IsDBNull(dr2("TIC_CF")) Then
                                    If Not String.IsNullOrEmpty(dr2("TIC_CF").ToString) Then
                                        pdfFormFields.SetField(pdfFieldPath & "L3[0]", FormatFixedAmount(dr2("TIC_CF").ToString))
                                    Else
                                        pdfFormFields.SetField(pdfFieldPath & "L3[0]", 0)
                                    End If
                                Else
                                    pdfFormFields.SetField(pdfFieldPath & "L3[0]", 0)
                                End If
                            Case 4
                                If Not IsDBNull(dr2("TIC_CF")) Then
                                    If Not String.IsNullOrEmpty(dr2("TIC_CF").ToString) Then
                                        pdfFormFields.SetField(pdfFieldPath & "L4[0]", FormatFixedAmount(dr2("TIC_CF").ToString))
                                    Else
                                        pdfFormFields.SetField(pdfFieldPath & "L4[0]", 0)
                                    End If
                                Else
                                    pdfFormFields.SetField(pdfFieldPath & "L4[0]", 0)
                                End If
                            Case 5
                                If Not IsDBNull(dr2("TIC_CF")) Then
                                    If Not String.IsNullOrEmpty(dr2("TIC_CF").ToString) Then
                                        pdfFormFields.SetField(pdfFieldPath & "L5[0]", FormatFixedAmount(dr2("TIC_CF").ToString))
                                    End If
                                End If
                            Case 6
                                If Not IsDBNull(dr2("TIC_CF")) Then
                                    If Not String.IsNullOrEmpty(dr2("TIC_CF").ToString) Then
                                        pdfFormFields.SetField(pdfFieldPath & "L6[0]", FormatFixedAmount(dr2("TIC_CF").ToString))
                                    End If
                                End If
                        End Select
                    End If
                End While
                dr2.Close()
            End If
            dr.Close()


            ' === Part M === '
            'M
            dr = datHandler.GetDataReader("SELECT * FROM PROFIT_LOSS_ACCOUNT WHERE PL_REF_NO = '" _
                                      & pdfForm.GetRefNo & "' AND PL_YA = '" & pdfForm.GetYA & "'" _
                                      & "and PL_MAINCOMPANY = '1' order by PL_KEY")

            'dr = datHandler.GetDataReader("Select BC_BUS_ENTITY, BC_CODE" _
            '                         & " From BUSINESS_SOURCE WHERE" _
            '                         & " ADJ_REF_NO= '" & pdfForm.GetRefNo & "' AND ADJ_YA = '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                dr.Close()
                dr = datHandler.GetDataReader("SELECT * FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" _
                                    & pdfForm.GetRefNo & "' AND [PL_YA] = '" & pdfForm.GetYA & "' and [PL_MAINCOMPANY] = '1'" _
                                    & "order by [PL_KEY] ")
            Else
                dr.Close()
                dr = datHandler.GetDataReader("SELECT * FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" _
                                    & pdfForm.GetRefNo & "' AND [PL_YA] = '" & pdfForm.GetYA & "' order by [PL_KEY] ")
            End If

            If dr.Read() Then
                dr2 = datHandler.GetDataReader("SELECT BC_BUS_ENTITY, BC_CODE, BC_COMPANY FROM [BUSINESS_SOURCE] WHERE [BC_KEY] = '" _
                                        & pdfForm.GetRefNo & "' AND [BC_YA] = '" & pdfForm.GetYA & "' AND " _
                                        & "[BC_BUSINESSSOURCE] = '" & Trim(dr("PL_MAIN_BUSINESS")) & "'")
                If dr2.Read() Then
                    ReDim strArray(1)
                    If Not IsDBNull(dr2("BC_BUS_ENTITY")) Then
                        If Not String.IsNullOrEmpty(dr2("BC_BUS_ENTITY")) Then
                            strArray = SplitText(dr2("BC_BUS_ENTITY").ToString, 28)
                            If Not String.IsNullOrEmpty(strArray(0)) Then
                                pdfFormFields.SetField(pdfFieldPath & "M1_1[0]", strArray(0).ToString.ToUpper)
                                pdfFormFields.SetField(pdfFieldPath & "M1_2[0]", strArray(1).ToString.ToUpper)
                            End If
                        End If
                    End If

                    If Not IsDBNull(dr2("BC_CODE")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M1A[0]", dr2("BC_CODE"))
                    End If

                End If
                dr2.Close()

                'M2

                If Not IsDBNull(dr("PL_SALES")) Then
                    pdfFormFields.SetField(pdfFieldPath & "M2[0]", FormatFixedAmount(dr("PL_SALES").ToString))
                End If

                'M3
                If Not IsDBNull(dr("PL_OP_STK")) Then
                    pdfFormFields.SetField(pdfFieldPath & "M3[0]", FormatFixedAmount(dr("PL_OP_STK").ToString))
                End If

                'M4
                If Not IsDBNull(dr("PL_PURCHASES_PRO_COST")) Then
                    pdfFormFields.SetField(pdfFieldPath & "M4[0]", FormatFixedAmount(dr("PL_PURCHASES_PRO_COST").ToString))
                End If

                'M5
                If Not IsDBNull(dr("PL_CLS_STK")) Then
                    pdfFormFields.SetField(pdfFieldPath & "M5[0]", FormatFixedAmount(dr("PL_CLS_STK").ToString))
                End If

                'M6
                If Not IsDBNull(dr("PL_COGS")) Then
                    pdfFormFields.SetField(pdfFieldPath & "M6[0]", FormatFixedAmount(dr("PL_COGS").ToString))
                End If

                'M7
                If Not IsDBNull(dr("PL_GROSS_PROFIT")) Then
                    If Not String.IsNullOrEmpty(dr("PL_GROSS_PROFIT")) Then
                        If CDbl(dr("PL_GROSS_PROFIT")) >= 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "M7_1[0]", "")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "M7_1[0]", "X")
                        End If
                        pdfFormFields.SetField(pdfFieldPath & "M7_2[0]", FormatFixedAmount(CDbl(dr("PL_GROSS_PROFIT")).ToString.Replace(".", "").Replace("-", "")))
                        dblTotalValue3 = CDbl(dr("PL_GROSS_PROFIT"))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M7_1[0]", "")
                    End If
                End If
                'M8
                dblTotalValue = 0
                dblTotalValue2 = 0

                dr2 = datHandler.GetDataReader("SELECT sum(cdbl([EXA_AMOUNT])) FROM [PL_INCOME_OTHERBUSINESS] WHERE [EXA_KEY] = " & dr("PL_KEY"))
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTotalValue = CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()
                dblTotalValue2 = dblTotalValue

                dr2 = datHandler.GetDataReader("SELECT [PL_KEY] FROM [PROFIT_LOSS_ACCOUNT]" _
                        & " WHERE [PL_REF_NO] = '" & pdfForm.GetRefNo & "' AND [PL_YA] = '" & pdfForm.GetYA & "' and [PL_KEY] <> " & dr("PL_KEY"))
                While dr2.Read()
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTotalValue2 = dblTotalValue2 + OtherSource_GrossProfitLoss(CLng(dr2.Item(0)), dr("PL_COMPANY"), dr("PL_MAIN_BUSINESS"), pdfForm.GetRefNo, pdfForm.GetYA)
                    End If
                End While
                dr2.Close()
                pdfFormFields.SetField(pdfFieldPath & "M8[0]", FormatFixedAmount(dblTotalValue2.ToString))

                'rmk
                dblTotalValue = dblTotalValue2 - dblTotalValue


                'M9
                'dblTotalValue2 = 0
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_INCOME_NONBUSINESS] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 47 ")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M9[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M9[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M9[0]", 0)
                End If
                dr2.Close()


                'M10
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_INCOME_NONBUSINESS] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 50 ")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M10[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M10[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M10[0]", 0)
                End If
                dr2.Close()


                ' === Cont. Page 7 === '

                'M11
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_INCOME_NONBUSINESS] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE between 48 and 49 ")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M11[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M11[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M11[0]", 0)
                End If
                dr2.Close()

                'M12
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_INCOME_NONBUSINESS] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 51 ")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTemp = dblTemp + CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()

                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_INCOME_NONTAXABLE] WHERE [EXA_KEY] = " & dr("PL_KEY"))
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTemp = dblTemp + CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()

                dblTotalValue2 = dblTotalValue2 + dblTemp
                dblTotalValue3 = dblTotalValue3 + dblTotalValue2 - dblTotalValue

                pdfFormFields.SetField(pdfFieldPath & "M12[0]", FormatFixedAmount(dblTemp.ToString))
                'M13
                pdfFormFields.SetField(pdfFieldPath & "M13[0]", FormatFixedAmount(dblTotalValue2.ToString))


                'M14
                dblTotalValue2 = 0
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 11")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M14[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M14[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M14[0]", 0)
                End If
                dr2.Close()


                'M15
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 12")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M15[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M15[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M15[0]", 0)
                End If
                dr2.Close()


                'M16
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 13")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M16[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M16[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M16[0]", 0)
                End If
                dr2.Close()


                'M17
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 14")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M17[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M17[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M17[0]", 0)
                End If
                dr2.Close()


                'M18
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 15")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M18[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M18[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M18[0]", 0)
                End If
                dr2.Close()


                'M19
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 16")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M19[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M19[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M19[0]", 0)
                End If
                dr2.Close()


                'M20
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 17")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M20[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M20[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M20[0]", 0)
                End If
                dr2.Close()


                'M21
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 52")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M21[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M21[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M21[0]", 0)
                End If
                dr2.Close()


                'M22
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND EXA_PLTYPE = 53")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        pdfFormFields.SetField(pdfFieldPath & "M22[0]", FormatFixedAmount(dr2.Item(0).ToString))
                        dblTotalValue2 = dblTotalValue2 + CDbl(dr2.Item(0))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M22[0]", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M22[0]", 0)
                End If
                dr2.Close()


                'M23 - M24
                dblTemp = 0
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND (EXA_PLTYPE between 18 and 20 or EXA_PLTYPE = 46)")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTemp = CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()

                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXP_NONALLOWLOSS] WHERE [EXA_KEY] = " & dr("PL_KEY"))
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTemp = dblTemp + CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()

                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXP_NONALLOWEXPEND] WHERE [EXA_KEY] = " & dr("PL_KEY"))
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTemp = dblTemp + CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()

                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXP_PERSONAL] WHERE [EXA_KEY] = " & dr("PL_KEY"))
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTemp = dblTemp + CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()

                dblTotalValue2 = dblTotalValue2 + dblTemp

                pdfFormFields.SetField(pdfFieldPath & "M23[0]", FormatFixedAmount(dblTemp.ToString))
                pdfFormFields.SetField(pdfFieldPath & "M24[0]", FormatFixedAmount(dblTotalValue2.ToString))


                'M25
                dblTotalValue3 = dblTotalValue3 - dblTotalValue2
                If dblTotalValue3 >= 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "M25_1[0]", "")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M25_1[0]", "X")
                End If
                pdfFormFields.SetField(pdfFieldPath & "M25_2[0]", FormatFixedAmount(dblTotalValue3.ToString.Replace(".", "").Replace("-", "")))


                'M26 Cal
                dblTemp = 0
                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXPENSES] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND [EXA_DEDUCTIBLE]='No'")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTemp = dblTemp + CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()

                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXP_NONALLOWEXPEND] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND [EXA_DEDUCTIBLE]='No'")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTemp = dblTemp + CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()

                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_EXP_PERSONAL] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND [EXA_DEDUCTIBLE]='No'")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTemp = dblTemp + CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()

                dr2 = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_PRODUCTION_COST] WHERE [EXA_KEY] = " & dr("PL_KEY") & " AND [EXA_DEDUCTIBLE]='No' and (EXA_PLTYPE = 43 or EXA_PLTYPE = 45)")
                If dr2.Read() Then
                    If Not IsDBNull(dr2.Item(0)) Then
                        dblTemp = dblTemp + CDbl(dr2.Item(0))
                    End If
                End If
                dr2.Close()

            Else
                pdfFormFields.SetField(pdfFieldPath & "M1A[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "M2[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M3[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M4[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M5[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M6[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M7_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "M7_2[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M8[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M9[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M10[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M11[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M12[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M13[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M14[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M15[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M16[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M17[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M18[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M19[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M20[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M21[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M22[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M23[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M24[0]", "0")
                pdfFormFields.SetField(pdfFieldPath & "M25_1[0]", "")
                pdfFormFields.SetField(pdfFieldPath & "M25_2[0]", "0")
            End If

            'M26
            pdfFormFields.SetField(pdfFieldPath & "M26[0]", FormatFixedAmount(dblTemp.ToString))
            dr.Close()



            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama8", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj8", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Page7()

        Dim pdfFieldPath As String
        Dim dr As OleDbDataReader = Nothing
        Dim dr2 As OleDbDataReader = Nothing
        Dim dblTotalValue As Double = 0
        Dim dblTotalValue2 As Double = 0
        Dim intCounter As Integer = 1
        Dim strArray(1) As String


        ' === Part M === '
        Try
            ' prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page9[0]."

            dr = datHandler.GetDataReader("SELECT * FROM PROFIT_LOSS_ACCOUNT WHERE PL_REF_NO = '" _
                                      & pdfForm.GetRefNo & "' AND PL_YA = '" & pdfForm.GetYA & "'" _
                                      & "and PL_MAINCOMPANY = '1' order by PL_KEY")

            'dr = datHandler.GetDataReader("Select BC_BUS_ENTITY, BC_CODE" _
            '                         & " From BUSINESS_SOURCE WHERE" _
            '                         & " ADJ_REF_NO= '" & pdfForm.GetRefNo & "' AND ADJ_YA = '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                dr.Close()
                dr = datHandler.GetDataReader("SELECT * FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" _
                                    & pdfForm.GetRefNo & "' AND [PL_YA] = '" & pdfForm.GetYA & "' and [PL_MAINCOMPANY] = '1'" _
                                    & "order by [PL_KEY] ")
            Else
                dr.Close()
                dr = datHandler.GetDataReader("SELECT * FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" _
                                    & pdfForm.GetRefNo & "' AND [PL_YA] = '" & pdfForm.GetYA & "' order by [PL_KEY] ")
            End If

            If dr.Read() Then

                dr2 = datHandler.GetDataReader("SELECT * FROM [BALANCE_SHEET] WHERE [BS_REF_NO] = '" & pdfForm.GetRefNo & "' AND [BS_YA] = '" & pdfForm.GetYA & "' and [BS_SOURCENO] = '" & Trim(dr("PL_MAIN_BUSINESS")) & "' order by BS_SOURCENO")
                If dr2.Read Then

                    'M27
                    If Not IsDBNull(dr2("BS_LAND")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M27", FormatFixedAmount(dr2("BS_LAND").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M27", 0)
                    End If

                    'M28
                    If Not IsDBNull(dr2("BS_MACHINERY")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M28", FormatFixedAmount(dr2("BS_MACHINERY").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M28", 0)
                    End If

                    'M29
                    If Not IsDBNull(dr2("BS_TRANSPORT")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M29", FormatFixedAmount(dr2("BS_TRANSPORT").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M29", 0)
                    End If

                    'M30
                    If Not IsDBNull(dr2("BS_OTH_FA")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M30", FormatFixedAmount(dr2("BS_OTH_FA").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M30", 0)
                    End If

                    'M31
                    If Not IsDBNull(dr2("BS_TOT_FA")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M31", FormatFixedAmount(dr2("BS_TOT_FA").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M31", 0)
                    End If

                    'M32
                    If Not IsDBNull(dr2("BS_INVESTMENT")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M32", FormatFixedAmount(dr2("BS_INVESTMENT").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M32", 0)
                    End If

                    'M33
                    If Not IsDBNull(dr2("BS_STOCK")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M33", FormatFixedAmount(dr2("BS_STOCK").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M33", 0)
                    End If

                    'M34
                    If Not IsDBNull(dr2("BS_TRADE_DEBTORS")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M34", FormatFixedAmount(dr2("BS_TRADE_DEBTORS").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M34", 0)
                    End If

                    'M35
                    If Not IsDBNull(dr2("BS_OTH_DEBTORS")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M35", FormatFixedAmount(dr2("BS_OTH_DEBTORS").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M35", 0)
                    End If

                    'M36
                    If Not IsDBNull(dr2("BS_CASH")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M36", FormatFixedAmount(dr2("BS_CASH").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M36", 0)
                    End If

                    'M37
                    If Not IsDBNull(dr2("BS_BANK")) Then
                        If CDbl(dr2("BS_BANK")) >= 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "M37_1", "")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "M37_1", "X")
                        End If
                        pdfFormFields.SetField(pdfFieldPath & "M37_2", FormatFixedAmount(FormatNumber(CDbl(dr2("BS_BANK")), 0).ToString.Replace("-", "")))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M37_1", "")
                        pdfFormFields.SetField(pdfFieldPath & "M37_2", 0)
                    End If

                    'M38
                    If Not IsDBNull(dr2("BS_OTH_CA")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M38", FormatFixedAmount(dr2("BS_OTH_CA").ToString))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M38", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M27", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M28", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M29", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M30", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M31", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M32", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M33", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M34", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M35", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M36", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M37_1", "")
                    pdfFormFields.SetField(pdfFieldPath & "M37_2", 0)
                    pdfFormFields.SetField(pdfFieldPath & "M38", 0)
                End If
                dr2.Close()
            Else
                pdfFormFields.SetField(pdfFieldPath & "M27", 0)
                pdfFormFields.SetField(pdfFieldPath & "M28", 0)
                pdfFormFields.SetField(pdfFieldPath & "M29", 0)
                pdfFormFields.SetField(pdfFieldPath & "M30", 0)
                pdfFormFields.SetField(pdfFieldPath & "M31", 0)
                pdfFormFields.SetField(pdfFieldPath & "M32", 0)
                pdfFormFields.SetField(pdfFieldPath & "M33", 0)
                pdfFormFields.SetField(pdfFieldPath & "M34", 0)
                pdfFormFields.SetField(pdfFieldPath & "M35", 0)
                pdfFormFields.SetField(pdfFieldPath & "M36", 0)
                pdfFormFields.SetField(pdfFieldPath & "M37_1", "")
                pdfFormFields.SetField(pdfFieldPath & "M37_2", 0)
                pdfFormFields.SetField(pdfFieldPath & "M38", 0)
            End If
            dr.Close()


            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama9", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj9", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Page8()

        Dim pdfFieldPath As String
        Dim strTempString As String = ""
        Dim dr As OleDbDataReader = Nothing
        Dim dr2 As OleDbDataReader = Nothing
        Dim dblTotalValue As Double = 0
        Dim dblTotalValue2 As Double = 0
        Dim intCounter As Integer = 1
        Dim strArray(1) As String


        Try
            'prmOledb(0) = New OleDbParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page10[0]."

            ' === Part M === '

            'dr = datHandler.GetDataReader("SELECT * FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" _
            '        & pdfForm.GetRefNo & "' AND [PL_YA] = '" & pdfForm.GetYA & "' order by [PL_KEY] ")

            dr = datHandler.GetDataReader("SELECT * FROM PROFIT_LOSS_ACCOUNT WHERE PL_REF_NO = '" _
                                      & pdfForm.GetRefNo & "' AND PL_YA = '" & pdfForm.GetYA & "'" _
                                      & "and PL_MAINCOMPANY = '1' order by PL_KEY")

            'dr = datHandler.GetDataReader("Select BC_BUS_ENTITY, BC_CODE" _
            '                         & " From BUSINESS_SOURCE WHERE" _
            '                         & " ADJ_REF_NO= '" & pdfForm.GetRefNo & "' AND ADJ_YA = '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                dr.Close()
                dr = datHandler.GetDataReader("SELECT * FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" _
                                    & pdfForm.GetRefNo & "' AND [PL_YA] = '" & pdfForm.GetYA & "' and [PL_MAINCOMPANY] = '1'" _
                                    & "order by [PL_KEY] ")
            Else
                dr.Close()
                dr = datHandler.GetDataReader("SELECT * FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" _
                                    & pdfForm.GetRefNo & "' AND [PL_YA] = '" & pdfForm.GetYA & "' order by [PL_KEY] ")
            End If

            If dr.Read() Then
                dr2 = datHandler.GetDataReader("SELECT * FROM [BALANCE_SHEET] WHERE [BS_REF_NO] = '" & pdfForm.GetRefNo & "' AND [BS_YA] = '" & pdfForm.GetYA & "' and [BS_SOURCENO] = '" & Trim(dr("PL_MAIN_BUSINESS")) & "' order by BS_SOURCENO")
                If dr2.Read Then

                    'M39
                    If Not IsDBNull(dr2("BS_TOT_CA")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M39", FormatFixedAmount(dr2("BS_TOT_CA").ToString))
                    End If
                    'M40
                    If Not IsDBNull(dr2("BS_TOT_ASSETS")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M40", FormatFixedAmount(dr2("BS_TOT_ASSETS").ToString))
                    End If
                    'M41
                    If Not IsDBNull(dr2("BS_LOAN")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M41", FormatFixedAmount(dr2("BS_LOAN").ToString))
                    End If
                    'M42
                    If Not IsDBNull(dr2("BS_TRADE_CR")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M42", FormatFixedAmount(dr2("BS_TRADE_CR").ToString))
                    End If
                    'M43
                    If Not IsDBNull(dr2("BS_OTHER_CR")) And Not IsDBNull(dr2("BS_OTH_LIAB")) And Not IsDBNull(dr2("BS_LT_LIAB")) Then
                        dblTotalValue = CDbl(dr2("BS_OTHER_CR")) + CDbl(dr2("BS_OTH_LIAB")) + CDbl(dr2("BS_LT_LIAB"))
                        pdfFormFields.SetField(pdfFieldPath & "M43", FormatFixedAmount(dblTotalValue.ToString))
                    End If
                    'M44
                    If Not IsDBNull(dr2("BS_TOT_LIAB")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M44", FormatFixedAmount(dr2("BS_TOT_LIAB").ToString))
                    End If
                    'M45
                    If Not IsDBNull(dr2("BS_CAPITALACCOUNT")) Then
                        pdfFormFields.SetField(pdfFieldPath & "M45", FormatFixedAmount(dr2("BS_CAPITALACCOUNT").ToString))
                    End If
                    'M46
                    If Not IsDBNull(dr2("BS_BROUGHT_FORWARD")) Then
                        If CDbl(dr2("BS_BROUGHT_FORWARD")) >= 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "M46_1", "")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "M46_1", "X")
                        End If
                        pdfFormFields.SetField(pdfFieldPath & "M46_2", FormatFixedAmount(CDbl(dr2("BS_BROUGHT_FORWARD")).ToString.Replace("-", "")))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M46_1", "")
                        pdfFormFields.SetField(pdfFieldPath & "M46_2", 0)
                    End If
                    'M47
                    If Not IsDBNull(dr2("BS_CY_PROFITLOSS")) Then
                        If CDbl(dr2("BS_CY_PROFITLOSS")) >= 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "M47_1", "")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "M47_1", "X")
                        End If
                        pdfFormFields.SetField(pdfFieldPath & "M47_2", FormatFixedAmount(CDbl(dr2("BS_CY_PROFITLOSS")).ToString.Replace("-", "")))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M47_1", "")
                        pdfFormFields.SetField(pdfFieldPath & "M47_2", 0)
                    End If
                    'M48
                    dblTotalValue = 0
                    If Not IsDBNull(dr2("BS_CAP_CONTRIBUTION")) And Not IsDBNull(dr2("BS_DRAWING")) Then
                        dblTotalValue = CDbl(dr2("BS_CAP_CONTRIBUTION")) - CDbl(dr2("BS_DRAWING"))
                        If dblTotalValue >= 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "M48_1", "")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "M48_1", "X")
                        End If
                        pdfFormFields.SetField(pdfFieldPath & "M48_2", FormatFixedAmount(dblTotalValue.ToString.Replace("-", "")))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M48_1", "")
                        pdfFormFields.SetField(pdfFieldPath & "M48_2", 0)
                    End If
                    'M49
                    If Not IsDBNull(dr2("BS_CARRIED_FORWARD")) Then
                        If CDbl(dr2("BS_CARRIED_FORWARD")) >= 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "M49_1", "")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "M49_1", "X")
                        End If
                        pdfFormFields.SetField(pdfFieldPath & "M49_2", FormatFixedAmount(CDbl(dr2("BS_CARRIED_FORWARD")).ToString.Replace("-", "")))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "M49_1", "")
                        pdfFormFields.SetField(pdfFieldPath & "M49_2", 0)
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "M39", "0")
                    pdfFormFields.SetField(pdfFieldPath & "M40", "0")
                    pdfFormFields.SetField(pdfFieldPath & "M41", "0")
                    pdfFormFields.SetField(pdfFieldPath & "M42", "0")
                    pdfFormFields.SetField(pdfFieldPath & "M43", "0")
                    pdfFormFields.SetField(pdfFieldPath & "M44", "0")
                    pdfFormFields.SetField(pdfFieldPath & "M45", "0")
                    pdfFormFields.SetField(pdfFieldPath & "M46_1", "")
                    pdfFormFields.SetField(pdfFieldPath & "M46_2", "0")
                    pdfFormFields.SetField(pdfFieldPath & "M47_1", "")
                    pdfFormFields.SetField(pdfFieldPath & "M47_2", "0")
                    pdfFormFields.SetField(pdfFieldPath & "M48_1", "")
                    pdfFormFields.SetField(pdfFieldPath & "M48_2", "0")
                    pdfFormFields.SetField(pdfFieldPath & "M49_1", "")
                    pdfFormFields.SetField(pdfFieldPath & "M49_2", "0")
                End If
                dr2.Close()
            Else
                pdfFormFields.SetField(pdfFieldPath & "M39", "0")
                pdfFormFields.SetField(pdfFieldPath & "M40", "0")
                pdfFormFields.SetField(pdfFieldPath & "M41", "0")
                pdfFormFields.SetField(pdfFieldPath & "M42", "0")
                pdfFormFields.SetField(pdfFieldPath & "M43", "0")
                pdfFormFields.SetField(pdfFieldPath & "M44", "0")
                pdfFormFields.SetField(pdfFieldPath & "M45", "0")
                pdfFormFields.SetField(pdfFieldPath & "M46_1", "")
                pdfFormFields.SetField(pdfFieldPath & "M46_2", "0")
                pdfFormFields.SetField(pdfFieldPath & "M47_1", "")
                pdfFormFields.SetField(pdfFieldPath & "M47_2", "0")
                pdfFormFields.SetField(pdfFieldPath & "M48_1", "")
                pdfFormFields.SetField(pdfFieldPath & "M48_2", "0")
                pdfFormFields.SetField(pdfFieldPath & "M49_1", "")
                pdfFormFields.SetField(pdfFieldPath & "M49_2", "0")
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


            ' === Part Nyata === ''
            ReDim strArray(0)
            dr = datHandler.GetDataReader("SELECT * FROM [TAXA_PROFILE] Where [TA_KEY] =" & pdfForm.GetTaxAgent)
            If dr.Read() Then
                If Not IsDBNull(dr("TA_CO_NAME")) Then
                    If Not String.IsNullOrEmpty(dr("TA_CO_NAME")) Then
                        strArray = SplitText(dr("TA_CO_NAME").ToString, 29)
                        pdfFormFields.SetField(pdfFieldPath & "NyataA", strArray(0).ToString.ToUpper)
                    End If
                End If
                pdfFormFields.SetField(pdfFieldPath & "Nyatab", FormatPhoneNumber("", dr("TA_TEL_NO").ToString, "", dr("TA_MOBILE").ToString))
                If Not IsDBNull(dr("TA_LICENSE")) Then
                    pdfFormFields.SetField(pdfFieldPath & "Nyatac", dr("TA_LICENSE"))
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
                        pdfFormFields.SetField(pdfFieldPath & "Nama10", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj10", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub Page9()

        Dim pdfFieldPath As String = ""
        Dim strTempString As String = ""
        Dim dr As OleDbDataReader = Nothing


        Try
            ' === Part Slip === '
            pdfFieldPath = pdfSubFormName & "Page11[0]."
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


            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama11", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj11", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub Page10()

        Dim pdfFieldPath As String = ""
        Dim strTempString As String = ""
        Dim dr As OleDbDataReader = Nothing


        Try
            ' === Part Slip === '
            pdfFieldPath = pdfSubFormName & "Page12[0]."
            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama12", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj12", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub Page11()

        Dim pdfFieldPath As String = ""
        Dim strTempString As String = ""
        Dim dr As OleDbDataReader = Nothing


        Try
            ' === Part Slip === '
            pdfFieldPath = pdfSubFormName & "Page13[0]."
            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama13", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj13", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub Page12()

        Dim pdfFieldPath As String = ""
        Dim strTempString As String = ""
        Dim dr As OleDbDataReader = Nothing


        Try
            ' === Part Slip === '
            pdfFieldPath = pdfSubFormName & "Page14[0]."
            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama14", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj14", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

#End Region

#Region "PNL Business Rules"

    Public Function OtherSource_GrossProfitLoss(ByVal cPNL_Key As Long, ByVal PNLCompany As String, ByVal PnLMainBusiness As String, ByVal strRefNo As String, ByVal strYA As String) As Double

        Dim i As Integer, J As Integer
        Dim arrSource() As String = Nothing
        Dim arrPNL(,) As Double
        Dim osTotal As Double
        Dim dr As OleDbDataReader = Nothing

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
            dr = datHandler.GetDataReader("SELECT sum([PL_AMOUNT]) FROM [PL_SALES] WHERE [PL_KEY] = " & cPNL_Key & " AND [PL_SOURCENO] = '" & Trim(arrSource(J)) & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then arrPNL(J, 0) = CDbl(dr.Item(0))
            End If
            dr.Close()
            ''*** Opening Stock
            dr = datHandler.GetDataReader("SELECT sum([PL_AMOUNT]) FROM [PL_OPENSTOCK] WHERE [PL_KEY] = " & cPNL_Key & " AND [PL_SOURCENO] = '" & Trim(arrSource(J)) & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then arrPNL(J, 1) = CDbl(dr.Item(0))
            End If
            dr.Close()
            ''*** Purchase
            dr = datHandler.GetDataReader("SELECT sum([PL_AMOUNT]) FROM [PL_PURCHASE] WHERE [PL_KEY] = " & cPNL_Key & " AND [PL_SOURCENO] = '" & Trim(arrSource(J)) & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then arrPNL(J, 2) = CDbl(dr.Item(0))
            End If
            dr.Close()
            ''*** Cost of Production
            dr = datHandler.GetDataReader("SELECT sum([EXA_AMOUNT]) FROM [PL_PRODUCTION_COST] WHERE [EXA_KEY] = " & cPNL_Key & " AND [EXA_SOURCENO] = '" & Trim(arrSource(J)) & "'")
            If dr.Read() Then
                If Not IsDBNull(dr.Item(0)) Then arrPNL(J, 3) = CDbl(dr.Item(0))
            End If
            dr.Close()
            ''*** Closing Stock
            dr = datHandler.GetDataReader("SELECT sum([PL_AMOUNT]) FROM [PL_CLOSESTOCK] WHERE [PL_KEY] = " & cPNL_Key & " AND [PL_SOURCENO] = '" & Trim(arrSource(J)) & "'")
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
                Case "Page10[0].NyataA[0]"
                    CheckFieldEmpty(de.Key.ToString, 29)
                Case "Page3[0].I_1[0]", _
                    "Page3[0].I_2[0]", _
                    "Page3[0].A14[0]", _
                    "Page3[0].A13[0]", _
                    "Page4[0].B1_1[0]", _
                    "Page4[0].B1_2[0]", _
                    "Page7[0].H1_1[0]", _
                    "Page7[0].H1_2[0]", _
                    "Page8[0].M1_1[0]", _
                    "Page8[0].M1_2[0]", _
                    "Page10[0].Akuan1_1[0]", _
                    "Page10[0].Akuan1_2[0]"
                    CheckFieldEmpty(de.Key.ToString, 28)
                Case "Page3[0].A8_1[0]", _
                    "Page3[0].A8_2[0]", _
                    "Page3[0].A8_3[0]", _
                    "Page3[0].A9_1[0]", _
                    "Page3[0].A9_2[0]", _
                    "Page3[0].A9_3[0]"
                    CheckFieldEmpty(de.Key.ToString, 26)
                Case "Page3[0].A10[0]", _
                    "Page10[0].Nyatab[0]"
                    CheckFieldEmpty(de.Key.ToString, 13)
                Case "Page3[0].A4[0]", _
                    "Page3[0].A9a[0]", _
                    "Page4[0].H3[0]", _
                    "Page4[0].H5[0]", _
                    "Page10[0].Akuan4[0]", _
                    "Page10[0].NyataTarikh[0]"
                    CheckFieldEmpty(de.Key.ToString, 8)
                Case "Page3[0].II_1[0]", _
                    "Page3[0].A1[0]", _
                    "Page4[0].B2_1[0]", _
                    "Page6[0].D11_1[0]", _
                    "Page6[0].D11_2[0]", _
                    "Page6[0].D11_3[0]", _
                    "Page6[0].D11a_1[0]", _
                    "Page6[0].D11a_3[0]", _
                    "Page6[0].D11b1_1[0]", _
                    "Page6[0].D11b1_3[0]", _
                    "Page6[0].D11b2_1[0]", _
                    "Page6[0].D11b2_3[0]", _
                    "Page6[0].D11c1_1[0]", _
                    "Page6[0].D11c1_3[0]", _
                    "Page6[0].D11c2_1[0]", _
                    "Page6[0].D11c2_3[0]", _
                    "Page6[0].E2b_2[0]"
                    pdfFormFields.SetField(de.Key.ToString, RTrim("--"))
                Case "Page3[0].A2[0]", _
                    "Page3[0].A3[0]", _
                    "Page3[0].A5[0]", _
                    "Page3[0].A6[0]", _
                    "Page3[0].A7[0]", _
                    "Page5[0].C35_1[0]", _
                    "Page8[0].M7_1[0]", _
                    "Page9[0].M25_1[0]", _
                    "Page9[0].M37_1[0]", _
                    "Page10[0].M46_1[0]", _
                    "Page10[0].M47_1[0]", _
                    "Page10[0].M48_1[0]", _
                    "Page10[0].M49_1[0]", _
                    "Page10[0].Akuan3[0]"
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
        Else
            strTemp = "0"
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
            strTemp = (strMobilePrefix & strMobile).Replace("-", "")
        End If
        Return strTemp

    End Function

    'NGOHCS B2010.2
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
