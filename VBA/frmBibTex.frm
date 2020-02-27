VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBibTex 
   Caption         =   "BibTex para MS Word - GitHub/ddauriol"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "frmBibTex.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBibTex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bnt_Analisar_Click()
    
    ' Barra de Status
    Label_Status.Caption = "Analisando arquivo"
    
    'Limpando a lista
    frmBibTex.ListBox_Tags.Clear
    
    'Limpar formularios
    frmBibTex.TextBox_Author.Value = ""
    frmBibTex.TextBox_Title.Value = ""
    frmBibTex.TextBox_Year.Value = ""
    frmBibTex.TextBox_number.Value = ""
    frmBibTex.TextBox_Publisher.Value = ""
    frmBibTex.TextBox_DOI.Value = ""
    frmBibTex.TextBox_Pages.Value = ""
    frmBibTex.TextBox_Journal.Value = ""
    frmBibTex.TextBox_Tag.Value = ""
    frmBibTex.ComboBox_Type.Value = ""
    frmBibTex.TextBox_City.Value = ""
    
    'Analisando
    ReadBibTex (Me.TextBoxInBibTex.Value)
    
    ' Barra de Status
    Label_Status.Caption = frmBibTex.ListBox_Tags.ListCount & " Citações encontradas"
    
End Sub

Private Sub bnt_InsertAll_Click()

    ' Barra de Status
    Label_Status.Caption = "Inserindo Citações"
    
    Dim ListaCitations As Integer
    Dim i As Integer
    ListaCitations = frmBibTex.ListBox_Tags.ListCount
    
    For i = 1 To ListaCitations
        Call GetElementsBibTex(BibTexAll(i), i)
        ' Barra de Status
        Label_Status.Caption = "Inserindo citação " & i + 1 & " de " & ListaCitations
        Call bnt_InsertOne_Click
        If Status = 1 Then
            frmBibTex.ListBox_Tags.List(i, 0) = "Ok"
        End If
    Next i
    
    ' Barra de Status
    Label_Status.Caption = ""
    
End Sub

Private Sub bnt_InsertOne_Click()

    Dim strXml As String
    Dim BibTexAuthor, BibTexTitle, BibTexYear, BibTexNumber, BibTexPublisher, BibTexDOI, _
    BibTexPages, BibTexJournal, BibTexTypeAtual, BibTexTagAtual, BibTexCity As String
    
    BibTexAuthor = frmBibTex.TextBox_Author.Value
    BibTexTitle = frmBibTex.TextBox_Title.Value
    BibTexYear = frmBibTex.TextBox_Year.Value
    BibTexNumber = frmBibTex.TextBox_number.Value
    BibTexPublisher = frmBibTex.TextBox_Publisher.Value
    BibTexDOI = frmBibTex.TextBox_DOI.Value
    BibTexPages = frmBibTex.TextBox_Pages.Value
    BibTexJournal = frmBibTex.TextBox_Journal.Value
    BibTexTagAtual = frmBibTex.TextBox_Tag.Value
    BibTexTypeAtual = frmBibTex.ComboBox_Type.Value
    BibTexCity = frmBibTex.TextBox_City.Value
    
    If (BibTexAuthor = "" Or BibTexTitle = "" Or BibTexTagAtual = "" Or BibTexTypeAtual = "") Then
        ErroItem = MsgBox("Alguns itens são obrigatórios", vbCritical, "Error")
        Status = 0
        GoTo SetStatus
    End If
       
    'Criando o XML WordCitation
    strXml = "<b:Source xmlns:b=""http://schemas.microsoft.com/office/word/2004/10/bibliography"">" & _
    "<b:Tag>" + BibTexTagAtual + "</b:Tag>" & _
    "<b:SourceType>" + BibTexTypeAtual + "</b:SourceType>" & _
    "<b:Author><b:Author>" & _
    "<b:NameList><b:Person><b:Last>Hezi</b:Last>" & _
    "<b:First>Mor</b:First></b:Person></b:NameList></b:Author>" & _
    "</b:Author>" & _
    "<b:Title>" + BibTexTitle + "</b:Title>" & _
    "<b:Year>" + BibTexYear + "</b:Year>" & _
    "<b:City>" + BibTexCity + "</b:City>" & _
    "<b:Publisher>" + BibTexPublisher + "</b:Publisher>" & _
    "<b:Volume>" + BibTexNumber + "</b:Volume>" & _
    "<b:Pages>" + BibTexPages + "</b:Pages>" & _
    "</b:Source>"
    
    AddBibSource (strXml)

SetStatus:
    If Status = 1 Then
        If IsNull(frmBibTex.ListBox_Tags) Then
        Else
            Dim i As Integer
            i = frmBibTex.ListBox_Tags.ListIndex
            frmBibTex.ListBox_Tags.List(i, 0) = "Ok"
        End If
        ' Barra de Status
        Label_Status.Caption = "Citação inserida com sucesso"
    Else
        ' Barra de Status
        Label_Status.Caption = "Falha ao inserir Citação"
    End If
    
End Sub

Private Sub ListBox_Tags_Click()

    Dim i As Integer
    i = frmBibTex.ListBox_Tags.ListIndex + 1
    Call GetElementsBibTex(BibTexAll(i), i)
    
End Sub

Private Sub UserForm_Initialize()

    'Load Source Type
    Dim SourceType As Variant
    SourceType = Array("Art", _
                        "ArticleInAPeriodical", _
                        "Book", _
                        "BookSection", _
                        "Case", _
                        "ConferenceProceedings", _
                        "DocumentFromInternetSite", _
                        "ElectronicSource", _
                        "Film", _
                        "InternetSite", _
                        "Interview", _
                        "JournalArticle", _
                        "Misc", _
                        "Patent", _
                        "Performance", _
                        "Report", _
                        "SoundRecording")
    Dim intI As Integer
    For intI = 1 To (UBound(SourceType) - LBound(SourceType))
        ComboBox_Type.AddItem SourceType(intI)
    Next intI
    
    ComboBox_Type.Value = SourceType(1)
    ListBox_Tags.ColumnCount = 2
    ListBox_Tags.ColumnWidths = "26;100"

End Sub
