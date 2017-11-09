Public Class clsBorangHandler

    Private strYA As String
    Private strFormType As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal strYA As String, ByVal strFormType As String)
        Me.strYA = strYA
        Me.strFormType = strFormType
    End Sub

    Private Function CreateBorangB() As Object
        Dim objBorangB As New Object
        Select Case (Me.strYA)
            Case "2008"
                objBorangB = New clsBorangB()
            Case "2009"
                objBorangB = New clsBorangB2009()
            Case "2010"
                objBorangB = New clsBorangB2010()
            Case "2011"
                objBorangB = New clsBorangB2011()
            Case "2012"
                objBorangB = New clsBorangB2012()
            Case "2013"
                objBorangB = New clsBorangB2013()
            Case "2014"
                objBorangB = New clsBorangB2014()
            Case "2015"
                objBorangB = New clsBorangB2015()
            Case "2016"
                objBorangB = New clsBorangB2016()
        End Select
        Return objBorangB
    End Function

    Private Function CreateBorangM() As Object
        Dim objBorangM As New Object
        Select Case (Me.strYA)
            Case "2008"
                objBorangM = New clsBorangM()
            Case "2009"
                objBorangM = New clsBorangM2009()
            Case "2010"
                objBorangM = New clsBorangM2010()
            Case "2011"
                objBorangM = New clsBorangM2011()
                'simkh 2012su2.2
            Case "2012"
                objBorangM = New clsBorangM2012()
                'simkh end
            Case "2014"
                objBorangM = New clsBorangM2014()
            Case "2015"
                objBorangM = New clsBorangM2015()
            Case "2016"
                objBorangM = New clsBorangM2016()
        End Select
        Return objBorangM
    End Function

    Private Function CreateBorangBE() As Object
        Dim objBorangBE As New Object
        Select Case (Me.strYA)
            Case "2008"
                objBorangBE = New clsBorangBE()
            Case "2009"
                objBorangBE = New clsBorangBE2009()
            Case "2010"
                objBorangBE = New clsBorangBE2010()
            Case "2011"
                objBorangBE = New clsBorangBE2011()
            Case "2012"
                objBorangBE = New clsBorangBE2012()
            Case "2013"
                objBorangBE = New clsBorangBE2013()
            Case "2014"
                objBorangBE = New clsBorangBE2014()
            Case "2015"
                objBorangBE = New clsBorangBE2015()
            Case "2016"
                objBorangBE = New clsBorangBE2016()
        End Select
        Return objBorangBE
    End Function

    Private Function CreateBorangP() As Object
        Dim objBorangP As New Object
        Select Case (Me.strYA)
            Case "2008"
                objBorangP = New clsBorangP()
            Case "2009"
                objBorangP = New clsBorangP2009()
            Case "2010"
                objBorangP = New clsBorangP2010()
            Case "2011"
                objBorangP = New clsBorangP2011()
            Case "2012"
                objBorangP = New clsBorangP2012()
            Case "2013"
                objBorangP = New clsBorangP2013()
            Case "2014"
                objBorangP = New clsBorangP2014()
            Case "2015"
                objBorangP = New clsBorangP2015()
            Case "2016"
                objBorangP = New clsBorangP2016()
        End Select
        Return objBorangP
    End Function

    Private Function CreateBorangCP30() As Object
        Dim objBorangCP30 As New Object
        Select Case (Me.strYA)
            Case "2008"
                objBorangCP30 = New clsBorangCP30()
            Case "2009"
                objBorangCP30 = New clsBorangCP302009()
            Case "2010"
                objBorangCP30 = New clsBorangCP302010()
            Case "2011"
                objBorangCP30 = New clsBorangCP302011()
            Case "2012"
                objBorangCP30 = New clsBorangCP302012()
            Case "2012"
                objBorangCP30 = New clsBorangCP302012()
            Case "2013"
                objBorangCP30 = New clsBorangCP302013()
            Case "2014"
                objBorangCP30 = New clsBorangCP302014()
            Case "2015"
                objBorangCP30 = New clsBorangCP302015()
            Case "2016"
                objBorangCP30 = New clsBorangCP302016()
        End Select
        Return objBorangCP30
    End Function

    Public Function CreateBorang() As Object
        Dim objBorang As New Object
        Select Case (Me.strFormType)
            Case "B"
                objBorang = CreateBorangB()
            Case "BE"
                objBorang = CreateBorangBE()
            Case "M"
                objBorang = CreateBorangM()
            Case "P"
                objBorang = CreateBorangP()
            Case "CP30"
                objBorang = CreateBorangCP30()
        End Select

        Return objBorang
    End Function

End Class
