Module MstructuresNEW
    Public Structure CommissionAgreementType
        Dim Name As String
        Dim NameAddressOneLine As String
        Dim xId As Integer
    End Structure
    Public Structure ContractorProposalsType
        Dim id As Integer
        Dim ContractorID As Integer
        Dim CustomerID As Integer
        Dim ComboCount As Integer
        Dim CreateDateString As String
        Dim ProposalFileName As String
        Dim CustomerName As String
        Dim CustomerAddress As String
        Dim ContractorDateOfProposal As String
        Dim ContractorQuoteNumber As String
        Dim CusNameAddress_Signature As String
        Dim ProposalDescription As String
        Dim LastModifiedDate As Long
    End Structure
    Public Structure XmasCardSentType
        Dim id As Integer
        Dim CusID As Integer
        Dim Year As Integer
        Dim Sent As Integer
    End Structure

    Public Structure ChristmasCardSentType
        Dim id As Integer
        Dim CusID As Integer
        Dim NameID As Integer
        Dim YearOfMailing As Integer
        Dim HasBeenMailed As Integer
    End Structure
    Public Structure ChristmasListType '11/13/2018
        Dim acctNum As Integer
        Dim NumberLines As Integer
        Dim FiveLines() As String 'actually 5 lines
        'Dim SevenLines() As String 'actually 7 lines
    End Structure
    Public Structure QuoteLineType ' 10/16/2018
        Dim AcctNum As Integer
        Dim QNumber As Integer
        Dim LineNumber As Integer
        Dim Quantity As Decimal
        Dim Description As String
        Dim CostPerItem As Decimal
        Dim TotalCostPerLine As Decimal
    End Structure
    Public Structure QuoteNumberType '[QuoteNumber] ' 10/16/2018
        Dim QNumber As Integer
        Dim QDate As Date
        Dim AcctNum As Integer
        Dim QAmount As Decimal
        Dim Qaccepted As Integer
        Dim AcceptedDate As Date
        Dim InvoiceCreated As Integer
        Dim SalesmanNumber As Integer
        Dim InvoiceAmount As Decimal
        Dim comment As String
        Dim PaymentTerms As String
        Dim FullName As String
        Dim LastModifiedDate As Long
        Dim CusNameAddress As String

    End Structure
    Public Structure InvoiceNumberType ' 10/16/2018
        'id	InvoiceNumber	QNumber	Amount	DueDate	SalesmanNumber	AmountPaid	PaidDate	comment
        Dim InvoiceNumber As Integer
        Dim Qnumber As Integer
        Dim Amount As Decimal
        Dim DueDate As Date
        Dim SalesmanNumber As Integer
        Dim AmountPaid As Decimal
        Dim PaidDate As Date
        Dim comment As String
    End Structure

    '==
    Structure FuType
        Dim id As Integer
        Dim fuDate As Date
        Dim CusId As Integer
        Dim msg As String
        Dim scUserID As Integer
        Dim checked As Integer
        Dim OrigDate As Date
        Dim compDate As Date
        Dim companyName As String
    End Structure
    Public Structure fileParts
        Dim FullName As String
        Dim name As String
        Dim ext As String
        Dim path As String
        Dim sortBy As String
    End Structure
    Structure TwoIntegersType
        Dim i1 As Integer
        Dim i2 As Integer
    End Structure
    Structure IntegerAndLongType
        Dim i As Integer
        Dim L As Long
    End Structure
    Structure CategoryDescriptionType
        Dim id As Integer
        Dim CategoryNumber As Integer
        Dim ItemNumber As Integer
        Dim CategoryDescription As String
    End Structure
    Structure StringAndTwoIntegersType
        Dim s As String
        Dim i1 As Integer
        Dim i2 As Integer
    End Structure
    Structure TwoStringsAndOneIntegerType
        Dim s1 As String
        Dim s2 As String
        Dim i1 As Integer
        'Sub New(ByVal ss1 As String, ByVal ss2 As String, ByVal ii1 As Integer)
        '    i1 = ii1
        '    s1 = ss1
        '    s2 = ss2
        'End Sub
    End Structure
    Structure FileCabinetRecordType
        Dim ID As Integer
        Dim Name As String
        Dim Note As String
        Dim DeleteFlag As Integer
    End Structure
    Structure BirthdayInfoType
        'Name               Birthday	days hence	acctNum	deleteflag	dtics           	ID
        Dim Name As String
        Dim Birthday As String
        Dim daysHence As Integer
        Dim acctNum As Integer
        Dim deleteFlag As Integer
        Dim dTics As Long
        Dim ID As Integer
    End Structure
    Structure TwoStringsType
        Dim s1 As String
        Dim s2 As String
    End Structure
    'Structure stringTwoIntegerType
    '    Dim str As String
    '    Dim ptr1 As Integer
    '    Dim ptr2 As Integer
    'End Structure
    'Structure StringPointerType
    '    Dim str As String
    '    Dim ptr As Integer
    'End Structure
    Structure ContactType

        <VBFixedString(20)> Dim FirstName As String
        <VBFixedString(20)> Dim LastName As String
        <VBFixedString(42)> Dim FullName As String

        Dim ID As Integer
        <VBFixedString(15)> Dim PhoneHome As String
        <VBFixedString(15)> Dim PhoneCell As String
        <VBFixedString(15)> Dim PhoneOffice As String

        <VBFixedString(30)> Dim Address1 As String
        <VBFixedString(30)> Dim Address2 As String
        <VBFixedString(30)> Dim City As String
        <VBFixedString(2)> Dim State As String
        <VBFixedString(20)> Dim Zip As String

    End Structure
    'Structure FileCabinetRecordType
    '    Dim RecNum As Integer
    '    <VBFixedString(92)> Dim FolderName As String
    '    Dim DeleteFlag As Integer
    'End Structure

    'Structure NamesRecordType
    '    Dim NamesAcctNum As Integer
    '    <VBFixedString(20)> Dim First As String
    '    <VBFixedString(20)> Dim Last As String
    '    <VBFixedString(20)> Dim Phone As String
    '    Dim DeleteFlag As Integer
    'End Structure
    'Structure KWUrecordType
    '    Dim KWUrecNum As Integer
    '    <VBFixedString(20)> Dim KW As String
    'End Structure
    'Structure LongIntegerType
    '    Dim L As Long
    '    Dim i As Integer
    'End Structure
    'Structure ApptRecTypeOLD '100 bybytes
    '    'Dim dTics As Date '8
    '    Dim dTics As Int64 'Date '8
    '    <VBFixedString(84)> Dim msg As String
    '    Dim DeleteFlag As Int32 '4
    '    Dim AcctNum As Int32 '4
    'End Structure
    'Structure ApptRecType '100 bybytes
    '    Dim dTics As Int64 'Date '8
    '    <VBFixedString(84)> Dim msg As String
    '    Dim AcctNum As Int32 '4
    '    Dim StrikeOut As Int16 '2 bytes 0 or 1
    '    Dim DeleteFlag As Int16 '2 bytes 0 or 1
    'End Structure
    Structure SubPartsType
        Dim FullFileName As String
        Dim Text As String
        Dim SubName As String
        Dim SortBy As String
        Dim xStart As Integer
    End Structure
    Public xStartTime As Long
    'Public StartTick As Long
    'Public EndTick As Long
    'Public Sub StartTime()
    '    StartTick = DateTime.Now.Ticks
    'End Sub
    'Public Sub EndTime()
    '    EndTick = DateTime.Now.Ticks
    'End Sub
    'Public Sub DiffTime()
    '    Dim L As Long
    '    L = EndTick - StartTick
    '    Debug.Print(CDec(L / 10000000))
    'End Sub
    '==
    Sub startTime()
        xStartTime = Now.Ticks
    End Sub
    Function fEndTime() As Decimal
        Dim x As Long = Now.Ticks
        Dim y As Decimal = x - xStartTime
        y /= 10000000
        Return y

    End Function
    Sub endTime()
        Dim x As Long = Now.Ticks
        Dim y As Decimal = x - xStartTime
        y /= 10000000
        Debug.Print(y.ToString)


    End Sub
    Sub initFiveStrings(ByRef c As ChristmasListType)
        Dim i As Integer
        ReDim c.FiveLines(4)
        For i = 0 To 4
            c.FiveLines(i) = ""
        Next
    End Sub

    'Sub init5Lines()
    '    Dim i As Integer
    '    For i = 0 To 4
    '        FiveLineArray(i) = ""
    '    Next
    'End Sub

End Module
