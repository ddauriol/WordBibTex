Attribute VB_Name = "BibTex"
Public BibTexAll() As String
Public TypeString() As String
Public BibTexTag() As String
Public Status As Integer
Sub WordBibTex()
    frmBibTex.Show
End Sub

Sub AddBibSource(strXml As String)
    
    On Error GoTo ErroInsert
    Application.Bibliography.Sources.Add strXml
    Status = 1
    Exit Sub
    
ErroInsert:
    Status = 0
    
End Sub

Public Sub ReadBibTex(TextString As String)
    
    Dim BibTexTemp As String
    Dim BibTex() As String
    Dim BibTexElements() As String
    Dim lNumBibTex, lNumElements As Integer
    Dim PosStart, PosEnd, LenString As Integer
    
    'Removendo excessos
    TextString = Replace(TextString, vbCrLf, " ")
    TextString = Replace(TextString, vbTab, " ")
    TextString = Replace(TextString, vbLf, " ")
    TextString = Replace(TextString, "  ", " ")
    TextString = Replace(TextString, ", ", ",")
    TextString = Replace(TextString, "  ", " ")
    TextString = Replace(TextString, "  ", " ")
    TextString = Replace(TextString, "{\'{a}}", "a")
    TextString = Replace(TextString, "{\'a}", "a")
    TextString = Replace(TextString, "{\" + Chr(34) + "{a}}", "a")
    TextString = Replace(TextString, "{\" + Chr(34) + "a}", "a")
    TextString = Replace(TextString, "{\~{a}}", "a")
    TextString = Replace(TextString, "{\~a}", "a")
    TextString = Replace(TextString, "{\'{e}}", "e")
    TextString = Replace(TextString, "{\'e}", "e")
    TextString = Replace(TextString, "{\" + Chr(34) + "{e}}", "e")
    TextString = Replace(TextString, "{\" + Chr(34) + "e}", "e")
    TextString = Replace(TextString, "{\^{e}}", "e")
    TextString = Replace(TextString, "{\^e}", "e")
    TextString = Replace(TextString, "{\'{i}}", "i&")
    TextString = Replace(TextString, "{\'i}", "i&")
    TextString = Replace(TextString, "{\" + Chr(34) + "{i}}", "i")
    TextString = Replace(TextString, "{\" + Chr(34) + "i}", "i")
    TextString = Replace(TextString, "{\'{o}}", "o")
    TextString = Replace(TextString, "{\'o}", "o")
    TextString = Replace(TextString, "{\" + Chr(34) + "{o}}", "o")
    TextString = Replace(TextString, "{\" + Chr(34) + "o}", "o")
    TextString = Replace(TextString, "{\~{o}}", "o")
    TextString = Replace(TextString, "{\~o}", "o")
    TextString = Replace(TextString, "{\^{o}}", "o")
    TextString = Replace(TextString, "{\^o}", "o")
    TextString = Replace(TextString, "{\'{u}}", "u")
    TextString = Replace(TextString, "{\'u}", "u")
    TextString = Replace(TextString, "{\" + Chr(34) + "{u}}", "u")
    TextString = Replace(TextString, "{\" + Chr(34) + "u}", "u")
    TextString = Replace(TextString, "{\c{c}}", "c")
    TextString = Replace(TextString, "{\cc}", "c")
    TextString = Replace(TextString, "{\AA}", "A")
    TextString = Replace(TextString, "{\%}", "%")
    TextString = Replace(TextString, "{\textperiodcentered}", "")
    TextString = Replace(TextString, "{\_}", "\")
    TextString = Replace(TextString, "{\textcopyright}", "")
    TextString = Replace(TextString, "{\&}", "&")
    
    'Identificando as citações
    BibTex = Split(TextString, "@")
    lNumBibTex = UBound(BibTex()) - LBound(BibTex())
    
    For i = 1 To lNumBibTex
        
        'Identificando o tipo do elemento
        PosEnd = InStr(BibTex(i), "{") - 1
        ReDim Preserve TypeString(i)
        TypeString(i) = Left(BibTex(i), PosEnd)
        
        'Identificando o TAG
        PosStart = PosEnd + 2
        PosEnd = InStr(BibTex(i), ",")
        LenString = PosEnd - PosStart
        ReDim Preserve BibTexTag(i)
        BibTexTag(i) = Mid(BibTex(i), PosStart, LenString)
        
        'Padronizando os dados
        BibTexTemp = Right(BibTex(i), Len(BibTex(i)) - PosEnd)
    
        BibTexTemp = Replace(BibTexTemp, "    ", " ")
        BibTexTemp = Replace(BibTexTemp, "   ", " ")
        BibTexTemp = Replace(BibTexTemp, "  ", " ")
        BibTexTemp = Replace(BibTexTemp, "  ", " ")
        
        BibTexTemp = Replace(BibTexTemp, "{{", "{")
        BibTexTemp = Replace(BibTexTemp, "{,", "{")
        BibTexTemp = Replace(BibTexTemp, "{ ", "{")
        
        BibTexTemp = Replace(BibTexTemp, "}}", "}")
        BibTexTemp = Replace(BibTexTemp, "},", "}")
        BibTexTemp = Replace(BibTexTemp, "} ", "}")
        
        BibTexTemp = Replace(BibTexTemp, "({", "(")
        BibTexTemp = Replace(BibTexTemp, ")}", ")")
        
        BibTexTemp = Replace(BibTexTemp, "{", "=")
        BibTexTemp = Replace(BibTexTemp, "}", "=")
        
        BibTexTemp = Replace(BibTexTemp, "= ", "=")
        BibTexTemp = Replace(BibTexTemp, "==", "=")
        BibTexTemp = Replace(BibTexTemp, " =", "=")
        
        'Arquivando as citações
        ReDim Preserve BibTexAll(i)
        BibTexAll(i) = BibTexTemp
        
        'Listando as citações
        frmBibTex.ListBox_Tags.AddItem "-"
        frmBibTex.ListBox_Tags.List(i - 1, 1) = BibTexTag(i)

        
    Next i
End Sub

Public Sub VerificaTipo(BibTexElements As String)
    
    If Replace(LCase(BibTexElements), " ", "") = "article" Then
        frmBibTex.ComboBox_Type.Value = "ArticleInAPeriodical"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "book" Then
        frmBibTex.ComboBox_Type.Value = "Book"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "booklet" Then
        frmBibTex.ComboBox_Type.Value = "BookSection"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "conference" Then
        frmBibTex.ComboBox_Type.Value = "JournalArticle"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "inbook" Then
        frmBibTex.ComboBox_Type.Value = "BookSection"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "incollection" Then
        frmBibTex.ComboBox_Type.Value = "BookSection"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "inproceedings" Then
        frmBibTex.ComboBox_Type.Value = "ArticleInAPeriodical"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "manual" Then
        frmBibTex.ComboBox_Type.Value = "Book"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "mastersthesis" Then
        frmBibTex.ComboBox_Type.Value = "Book"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "misc" Then
        frmBibTex.ComboBox_Type.Value = "Report"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "phdthesis" Then
        frmBibTex.ComboBox_Type.Value = "Book"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "proceedings" Then
        frmBibTex.ComboBox_Type.Value = "Report"
    ElseIf Replace(LCase(BibTexElements), " ", "") = "techreport" Then
        frmBibTex.ComboBox_Type.Value = "Report"
    Else:
        frmBibTex.ComboBox_Type.Value = "ArticleInAPeriodical"
    End If

End Sub

Public Sub GetElementsBibTex(BibTexTemp As String, i As Integer)

    Dim lNumElements As Integer
    Dim BibTexElements() As String
    
    'Criando os elementos
    BibTexElements = Split(BibTexTemp, "=")
    lNumElements = UBound(BibTexElements()) - LBound(BibTexElements())
       
    'Configurando a TAG no formulário
    frmBibTex.TextBox_Tag.Value = BibTexTag(i)
    
    'Configurando o tipo no formulário
    VerificaTipo (TypeString(i))
    BibTexType = frmBibTex.ComboBox_Type.Value
    
    For j = 0 To lNumElements
        If BibTexElements(j) <> "" Then
            If Replace(LCase(BibTexElements(j)), " ", "") = "author" Then
                frmBibTex.TextBox_Author.Value = BibTexElements(j + 1)
            ElseIf Replace(LCase(BibTexElements(j)), " ", "") = "title" Then
                frmBibTex.TextBox_Title.Value = BibTexElements(j + 1)
            ElseIf Replace(LCase(BibTexElements(j)), " ", "") = "year" Then
                frmBibTex.TextBox_Year.Value = BibTexElements(j + 1)
            ElseIf (Replace(LCase(BibTexElements(j)), " ", "") = "number" Or Replace(LCase(BibTexElements(j)), " ", "") = "volume") Then
                frmBibTex.TextBox_number.Value = BibTexElements(j + 1)
            ElseIf (Replace(LCase(BibTexElements(j)), " ", "") = "publisher" Or Replace(LCase(BibTexElements(j)), " ", "") = "editor") Then
                frmBibTex.TextBox_Publisher.Value = BibTexElements(j + 1)
            ElseIf (Replace(LCase(BibTexElements(j)), " ", "") = "doi" Or Replace(LCase(BibTexElements(j)), " ", "") = "isbn") Then
                frmBibTex.TextBox_DOI.Value = BibTexElements(j + 1)
            ElseIf Replace(LCase(BibTexElements(j)), " ", "") = "pages" Then
                frmBibTex.TextBox_Pages.Value = BibTexElements(j + 1)
            ElseIf (Replace(LCase(BibTexElements(j)), " ", "") = "journal" Or Replace(LCase(BibTexElements(j)), " ", "") = "booktitle") Then
                frmBibTex.TextBox_Journal.Value = BibTexElements(j + 1)
            End If
        End If
    Next j
        
End Sub
