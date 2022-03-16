Class MainWindow
    Private Sub Button1_Click(sender As Object, e As RoutedEventArgs) Handles Button1.Click

        ' Need a form with following fields:
        ' 1) TextBox1 to hold a URL to be scrubbed
        ' 2) TextBox2 to hold the resulting page scrub information
        ' 3) Button1 to start the scrub
        ' 4) Button2 to stop further processing or finalize


        Dim My_URL As String
        Dim My_Obj As Object
        Dim My_Var As String
        '        Dim s As String
        ' Dim My_Quote As String
        ' Dim StrQuoteAuthor As String
        Dim IntQuoteStarts As Integer
        ' Dim IntQuoteEnds As Integer
        Dim QuotesCount As Integer
        Dim AAuthorQuote(10, 2) As String
        Dim iNumQuotes As Integer
        '     Dim QTheme(4) As Collection
        Dim QTheme As New List(Of String) From {"BR", "NA", "AR", "FU", "LO"}
        Dim WhichTheme As Integer


        Debug.Print(QTheme(0))
        WhichTheme = ComboBox1.SelectedIndex

        TextBox1.Text = "http://feeds.feedburner.com/brainyquote/QUOTE" & QTheme(WhichTheme)

        My_URL = TextBox1.Text

        Debug.Print(My_URL)

        ' Button2.Text = "Done"

        ' Code modified to obtain specifically the 1st. quote on an RSS feed at: http://feeds.feedburner.com/brainyquote/QUOTEBR
        My_Obj = (CreateObject("MSXML2.XMLHTTP"))
        My_Obj.Open("GET", My_URL, False)
        My_Obj.send()
        My_Var = My_Obj.responsetext

        ' For debugging:
        ' Clipboard.SetText(My_Var)

        '       If InStr(1, My_Var, "<description>""") > 0 Then

        ' Get Author, should the first <item>, right after a <title> tag

        IntQuoteStarts = InStr(1, My_Var, "<feedburner:info ") ' find the tag of the 1st quote to get the show on the road

        My_Var = Mid(My_Var, IntQuoteStarts, Len(My_Var) - IntQuoteStarts)

        ' Get number of quotes found in the page
        GetQuotes(iNumQuotes, My_Var, "<title>", AAuthorQuote, )
        QuotesCount = iNumQuotes

        ' This section used to be the main code to obtain the 1st quote, replaced by the iterative function GetQuotes
        'right after this tag is the author name, just before /title tag
        '        IntQuoteStarts = InStr(IntQuoteStarts, My_Var, "<title>") + 15
        '        IntQuoteEnds = InStr(IntQuoteStarts, My_Var, "</title>") - IntQuoteStarts - 1
        '        StrQuoteAuthor = Mid(My_Var, IntQuoteStarts + 1, IntQuoteEnds)

        ' Get the 1st daily quote, should the first string, right after a <description>" tag [Notice triple quotes]
        '        IntQuoteStarts = InStr(1, My_Var, "<description>""") + 13
        '        IntQuoteEnds = InStr(IntQuoteStarts, My_Var, "</description>") - IntQuoteStarts
        '        My_Quote = Mid(My_Var, IntQuoteStarts, IntQuoteEnds)

        '        TextBox2.Text = My_Quote & " - " & StrQuoteAuthor

        '        GetQuote(My_Var, My_Quote)
        TextBox2.Text = ""
        For i = 1 To iNumQuotes
            TextBox2.Text = TextBox2.Text & AAuthorQuote(i, 0) & "-" & AAuthorQuote(i, 1) & vbCrLf
        Next

        My_Obj = Nothing

    End Sub
    Private Sub DoEvents()
        Throw New NotImplementedException
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click


        End
    End Sub

    ' Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    'Button2.Text = "Cancel"
    '    TextBox1.Text = "http://feeds.feedburner.com/brainyquote/QUOTEBR"
    ' End Sub
    Private Sub GetQuote(ByVal PageContent, ByRef TheQuote)
        Clipboard.SetText(PageContent)

    End Sub
    Public Sub GetQuotes(ByRef iQfound As Integer, ByVal OrigString As String,
      ByVal Chars As String, ByRef AReceive _
      As Object, Optional ByVal CaseSensitive As Boolean = False)

        '**********************************************
        'PURPOSE: Returns Number of occurrences of a character or
        'or a character sequencence within a string

        'PARAMETERS:
        'iQfound: number of quotes found
        'OrigString: String to Search in
        'Chars: Character(s) to search for
        'AReceive: 2 dimensional table to hold 2 elements of the quote found
        'CaseSensitive (Optional): Do a case sensitive search
        'Defaults to false

        'RETURNS:
        'Number of Occurrences of Chars in OrigString
        '2 dimensional array

        'EXAMPLES:
        'Debug.Print GetQuotes("FreeVBCode.com", "E") -- returns 3
        'Debug.Print GetQuotes("FreeVBCode.com", "E", True) -- returns 0
        'Debug.Print GetQuotes("FreeVBCode.com", "co") -- returns 2
        ''**********************************************
        'USAGE:
        '           Dim AAuthorQuote(10, 2) As String
        ' 	        GetQuotes(iNumQuotes, My_Var, "<title>", AAuthorQuote, )

        Dim lLen As Long
        Dim lCharLen As Long
        Dim lAns As Integer
        Dim sInput As String
        Dim sChar As String
        Dim iCtr As Integer
        Dim lEndOfLoop As Long
        Dim bytCompareType As Byte
        Dim iQuoteStarts As Integer
        Dim iQuoteEnds As Integer
        Dim sQuoteAuthor As String
        Dim sAuthorQuote(20, 2) As String
        Dim sTheQuote As String

        sInput = OrigString
        If sInput = "" Then Exit Sub
        lLen = Len(sInput)
        lCharLen = Len(Chars)
        lEndOfLoop = (lLen - lCharLen) + 1
        bytCompareType = IIf(CaseSensitive, vbBinaryCompare,
           vbTextCompare)

        For iCtr = 1 To lEndOfLoop
            sChar = Mid(sInput, iCtr, lCharLen)
            If StrComp(sChar, Chars, bytCompareType) = 0 Then
                lAns = lAns + 1
                'right after this tag is the author name, just before /title tag
                iQuoteStarts = InStr(iCtr, OrigString, "<title>") + 6
                iQuoteEnds = InStr(iQuoteStarts, OrigString, "</title>") - iQuoteStarts - 1
                sQuoteAuthor = Mid(OrigString, iQuoteStarts + 1, iQuoteEnds)

                ' Get the 1st daily quote, should the first string, right after a <description>" tag [Notice triple quotes]
                iQuoteStarts = InStr(iCtr, OrigString, "<description>""") + 13
                iQuoteEnds = InStr(iQuoteStarts, OrigString, "</description>") - iQuoteStarts
                sTheQuote = Mid(OrigString, iQuoteStarts, iQuoteEnds)

                ' Build the array
                sAuthorQuote(lAns, 0) = sTheQuote
                sAuthorQuote(lAns, 1) = sQuoteAuthor

            End If

        Next

        iQfound = lAns
        AReceive = sAuthorQuote

    End Sub

    Private Sub QODWindow_Activated(sender As Object, e As EventArgs) Handles QODWindow.Activated

        TextBox1.Text = "http://feeds.feedburner.com/brainyquote/QUOTE"
        ComboBox1.TabIndex = 1
        Button1.TabIndex = 2
        Button2.TabIndex = 3
        TextBox1.IsTabStop = False
        TextBox2.IsTabStop = False

    End Sub

    Private Sub Button1_ContextMenuClosing(sender As Object, e As ContextMenuEventArgs) Handles Button1.ContextMenuClosing

    End Sub

    Private Sub MainWindow_Activated(sender As Object, e As EventArgs) Handles Me.Activated

    End Sub

    Private Sub MainWindow_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered

    End Sub

    Private Sub ComboBox1_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs) Handles ComboBox1.ContextMenuOpening
    End Sub

    Private Sub ComboBox1_DropDownOpened(sender As Object, e As EventArgs) Handles ComboBox1.DropDownOpened

    End Sub

    Private Sub ComboBox1_ContextMenuClosing(sender As Object, e As ContextMenuEventArgs) Handles ComboBox1.ContextMenuClosing

    End Sub

    Private Sub ComboBox1_DropDownClosed(sender As Object, e As EventArgs) Handles ComboBox1.DropDownClosed

    End Sub

    Private Sub ComboBox1_Initialized(sender As Object, e As EventArgs) Handles ComboBox1.Initialized

        ComboBox1.Items.Add("Brainy")
        ComboBox1.Items.Add("Nature")
        ComboBox1.Items.Add("Art")
        ComboBox1.Items.Add("Funny")
        ComboBox1.Items.Add("Love")

    End Sub

    Private Sub MainWindow_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized

    End Sub

    Private Sub ComboBox1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboBox1.SelectionChanged

    End Sub

    Private Sub XGogetQOD(sender As Object, e As RoutedEventArgs) Handles ComboBox1.SelectionChanged
        ' Need a form with following fields:
        ' 1) TextBox1 to hold a URL to be scrubbed
        ' 2) TextBox2 to hold the resulting page scrub information
        ' 3) Button1 to start the scrub
        ' 4) Button2 to stop further processing or finalize


        Dim My_URL As String
        Dim My_Obj As Object
        Dim My_Var As String
        '        Dim s As String
        ' Dim My_Quote As String
        ' Dim StrQuoteAuthor As String
        Dim IntQuoteStarts As Integer
        ' Dim IntQuoteEnds As Integer
        Dim QuotesCount As Integer
        Dim AAuthorQuote(10, 2) As String
        Dim iNumQuotes As Integer
        '     Dim QTheme(4) As Collection
        Dim QTheme As New List(Of String) From {"BR", "NA", "AR", "FU", "LO"}
        Dim WhichTheme As Integer


        Debug.Print(QTheme(0))
        WhichTheme = ComboBox1.SelectedIndex

        TextBox1.Text = "http://feeds.feedburner.com/brainyquote/QUOTE" & QTheme(WhichTheme)

        My_URL = TextBox1.Text

        Debug.Print(My_URL)

        ' Button2.Text = "Done"

        ' Code modified to obtain specifically the 1st. quote on an RSS feed at: http://feeds.feedburner.com/brainyquote/QUOTEBR
        My_Obj = (CreateObject("MSXML2.XMLHTTP"))
        My_Obj.Open("GET", My_URL, False)
        My_Obj.send()
        My_Var = My_Obj.responsetext

        ' For debugging:
        ' Clipboard.SetText(My_Var)

        '       If InStr(1, My_Var, "<description>""") > 0 Then

        ' Get Author, should the first <item>, right after a <title> tag

        IntQuoteStarts = InStr(1, My_Var, "<feedburner:info ") ' find the tag of the 1st quote to get the show on the road

        My_Var = Mid(My_Var, IntQuoteStarts, Len(My_Var) - IntQuoteStarts)

        ' Get number of quotes found in the page
        GetQuotes(iNumQuotes, My_Var, "<title>", AAuthorQuote, )
        QuotesCount = iNumQuotes

        ' This section used to be the main code to obtain the 1st quote, replaced by the iterative function GetQuotes
        'right after this tag is the author name, just before /title tag
        '        IntQuoteStarts = InStr(IntQuoteStarts, My_Var, "<title>") + 15
        '        IntQuoteEnds = InStr(IntQuoteStarts, My_Var, "</title>") - IntQuoteStarts - 1
        '        StrQuoteAuthor = Mid(My_Var, IntQuoteStarts + 1, IntQuoteEnds)

        ' Get the 1st daily quote, should the first string, right after a <description>" tag [Notice triple quotes]
        '        IntQuoteStarts = InStr(1, My_Var, "<description>""") + 13
        '        IntQuoteEnds = InStr(IntQuoteStarts, My_Var, "</description>") - IntQuoteStarts
        '        My_Quote = Mid(My_Var, IntQuoteStarts, IntQuoteEnds)

        '        TextBox2.Text = My_Quote & " - " & StrQuoteAuthor

        '        GetQuote(My_Var, My_Quote)
        TextBox2.Text = ""
        For i = 1 To iNumQuotes
            TextBox2.Text = TextBox2.Text & AAuthorQuote(i, 0) & "-" & AAuthorQuote(i, 1) & vbCrLf
        Next

        My_Obj = Nothing
    End Sub
End Class

