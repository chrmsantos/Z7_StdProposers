Attribute VB_Name = "Mod_13_EmentaCoherence"
Option Explicit

' Mod_13_EmentaCoherence.bas
' =============================================================================
' Z7_STDPROPOSERS - Coerencia entre EMENTA e CORPO da propositura
' =============================================================================
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' =============================================================================
' Compara a ementa com o corpo (dispositivo) da propositura e avisa quando:
'   1. ENDERECO - logradouro citado na ementa nao aparece no corpo
'                 (ex.: ementa "Rua do Cacau", corpo "Rua do Milho")
'   2. NUMERO   - numero (2+ digitos) citado na ementa nao aparece no corpo
'   3. ASSUNTO  - ementa e corpo tratam de assuntos diferentes
'                 (ex.: ementa "vagas de estacionamento", corpo "tapa-buracos")
'
' O modulo e SOMENTE LEITURA: nunca altera o documento (seguro dentro do
' UndoRecord do pipeline). Nao usa IA nem rede: comparacao local e instantanea.
' Reaproveita: NormalizeForComparison e LevenshteinDistance (Mod_01),
' LogMessage (Mod_05), GetEmentaRange e GetCorpoRange (Mod_04).
'
' USO:
'   - Automatico: Mod_04 chama CheckEmentaCoherence apos BuildParagraphCache e
'     ShowEmentaCoherenceWarning ao final do pipeline.
'   - Manual: macro VerificarCoerenciaEmenta (documento ativo).
'   - Autoteste: macro TestarCoerenciaEmenta (nao precisa de documento).
' =============================================================================

' Mensagem pendente (preenchida por CheckEmentaCoherence, exibida ao final)
Public echLastWarning As String

Private Const ECH_ENABLED As Boolean = True
Private Const ECH_MAX_CHARS As Long = 20000
Private Const ECH_CAT_COUNT As Long = 9
Private Const ECH_MAX_NAME_TOKENS As Long = 6
Private Const ECH_TITLE As String = "Z7_STDPROPOSERS - Verificar Ementa"

' Cache de radicais das categorias (preenchido uma unica vez por ECH_EnsureStems,
' evitando re-splitar ECH_CatStems a cada token analisado)
Private echStems(1 To ECH_CAT_COUNT) As Variant
Private echStemsReady As Boolean

'================================================================================
' PONTOS DE ENTRADA
'================================================================================

' Verifica o documento. Retorna True se nao ha divergencia (ou se nao foi
' possivel verificar). Em caso de divergencia retorna False, registra no log e
' guarda a mensagem em echLastWarning. Requer estrutura ja identificada.
Public Function CheckEmentaCoherence(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    CheckEmentaCoherence = True
    echLastWarning = ""

    If Not ECH_ENABLED Then Exit Function
    If doc Is Nothing Then Exit Function

    Dim rngE As Range
    Dim rngC As Range
    Set rngE = GetEmentaRange(doc)
    Set rngC = GetCorpoRange(doc)

    If rngE Is Nothing Or rngC Is Nothing Then
        LogMessage "COERENCIA EMENTA: ementa ou corpo nao identificados - verificacao ignorada", LOG_LEVEL_INFO
        Exit Function
    End If

    Dim issues As String
    issues = ECH_CompareTexts(rngE.Text, rngC.Text)

    If Len(issues) = 0 Then
        LogMessage "COERENCIA EMENTA: nenhuma divergencia detectada", LOG_LEVEL_INFO
        Exit Function
    End If

    LogMessage "COERENCIA EMENTA: divergencia(s) detectada(s)" & vbCrLf & issues, LOG_LEVEL_WARNING

    echLastWarning = "Possiveis divergencias entre a EMENTA e o CORPO do texto:" & _
                     vbCrLf & vbCrLf & issues & vbCrLf & vbCrLf & _
                     "Confira o documento antes de protocolar."
    CheckEmentaCoherence = False
    Exit Function

ErrorHandler:
    ' Verificacao auxiliar: falha aqui nunca deve interromper a padronizacao
    LogMessage "COERENCIA EMENTA: erro na verificacao - " & Err.Description, LOG_LEVEL_WARNING
    echLastWarning = ""
    CheckEmentaCoherence = True
End Function

' Exibe (uma unica vez) o aviso pendente. Chamar ao FINAL do pipeline.
Public Sub ShowEmentaCoherenceWarning()
    If Len(echLastWarning) = 0 Then Exit Sub

    Dim msg As String
    msg = echLastWarning
    echLastWarning = ""

    MsgBox msg, vbExclamation, ECH_TITLE
End Sub

' Macro manual: verifica o documento ativo sem padronizar.
' Obs.: identifica a estrutura antes (IA com fallback para heuristica).
Public Sub VerificarCoerenciaEmenta()
    Dim doc As Document

    If Documents.count = 0 Then
        MsgBox "Nenhum documento esta aberto.", vbExclamation, ECH_TITLE
        Exit Sub
    End If

    Set doc = ActiveDocument
    IdentifyDocumentStructure doc

    If CheckEmentaCoherence(doc) Then
        MsgBox "Nenhuma divergencia detectada entre a ementa e o corpo.", vbInformation, ECH_TITLE
    Else
        ShowEmentaCoherenceWarning
    End If
End Sub

'================================================================================
' COMPARACAO (funcao pura - testavel sem documento)
'================================================================================

' Retorna as divergencias (uma por linha) ou "" se estiver coerente.
Public Function ECH_CompareTexts(ByVal ementaText As String, ByVal corpoText As String) As String
    Dim eNames As Collection, eKinds As Collection, eNums As Collection
    Dim cNames As Collection, cKinds As Collection, cNums As Collection
    Dim eMask As Long, cMask As Long
    Dim i As Long, j As Long
    Dim found As Boolean
    Dim lines As String
    Dim missing As String

    ECH_Analyze ementaText, eNames, eKinds, eNums, eMask
    ECH_Analyze corpoText, cNames, cKinds, cNums, cMask

    ' ---- 1. Enderecos ------------------------------------------------------
    For i = 1 To eNames.count
        found = False
        For j = 1 To cNames.count
            If ECH_SameName(CStr(eNames(i)), CStr(cNames(j))) Then
                found = True
                Exit For
            End If
        Next j

        If Not found Then
            lines = lines & "- ENDERECO: a ementa cita """ & eKinds(i) & " " & eNames(i) & _
                    """, que nao aparece no corpo"
            If cNames.count > 0 Then
                lines = lines & " (o corpo cita: " & ECH_ListStreets(cNames, cKinds) & ")."
            Else
                lines = lines & " (o corpo nao cita nenhum logradouro)."
            End If
            lines = lines & vbCrLf
        End If
    Next i

    ' ---- 2. Numeros --------------------------------------------------------
    For i = 1 To eNums.count
        If Not ECH_ColContains(cNums, CStr(eNums(i))) Then
            If Len(missing) > 0 Then missing = missing & ", "
            missing = missing & eNums(i)
        End If
    Next i

    If Len(missing) > 0 Then
        lines = lines & "- NUMERO: a ementa cita " & missing & ", que nao aparece no corpo"
        If cNums.count > 0 Then
            lines = lines & " (o corpo cita: " & ECH_ListPlain(cNums) & ")."
        Else
            lines = lines & " (o corpo nao cita numeros)."
        End If
        lines = lines & vbCrLf
    End If

    ' ---- 3. Assunto --------------------------------------------------------
    If eMask <> 0 And cMask <> 0 Then
        If (eMask And cMask) = 0 Then
            lines = lines & "- ASSUNTO: a ementa trata de " & ECH_MaskToText(eMask) & _
                    ", mas o corpo trata de " & ECH_MaskToText(cMask) & "." & vbCrLf
        End If
    End If

    If Len(lines) > 2 Then lines = Left$(lines, Len(lines) - 2) ' remove ultimo vbCrLf
    ECH_CompareTexts = lines
End Function

'================================================================================
' ANALISE DO TEXTO: logradouros, numeros e assunto
'================================================================================

Private Sub ECH_Analyze(ByVal text As String, _
                        ByRef stNames As Collection, _
                        ByRef stKinds As Collection, _
                        ByRef numList As Collection, _
                        ByRef catMask As Long)
    Dim toks() As String
    Dim isName() As Boolean
    Dim n As Long
    Dim i As Long, j As Long, k As Long, cnt As Long
    Dim nm As String
    Dim w As String

    Set stNames = New Collection
    Set stKinds = New Collection
    Set numList = New Collection
    catMask = 0

    n = ECH_Tokenize(text, toks)
    If n = 0 Then Exit Sub

    ReDim isName(1 To n)

    ' Logradouros: tipo (rua, avenida...) + nome ate um delimitador
    i = 1
    Do While i <= n
        If ECH_IsStreetType(toks(i)) Then
            j = i + 1
            cnt = 0
            nm = ""
            Do While j <= n
                If cnt >= ECH_MAX_NAME_TOKENS Then Exit Do
                w = toks(j)
                If w = "|" Then Exit Do
                If ECH_IsStopWord(w) Then Exit Do
                If ECH_IsDigits(w) Then Exit Do
                cnt = cnt + 1
                If Not ECH_IsArticle(w) Then
                    If Len(nm) > 0 Then nm = nm & " "
                    nm = nm & w
                End If
                j = j + 1
            Loop

            If Len(nm) > 0 Then
                stNames.Add nm
                stKinds.Add toks(i)
                For k = i + 1 To j - 1
                    isName(k) = True
                Next k
            End If
            i = j
        Else
            i = i + 1
        End If
    Loop

    ' Numeros (2+ digitos, exceto "108" do Art. 108) e assunto
    ' (palavras que fazem parte do nome de rua nao entram no assunto)
    For i = 1 To n
        w = toks(i)
        If ECH_IsDigits(w) Then
            If Len(w) >= 2 And w <> "108" Then
                If Not ECH_ColContains(numList, w) Then numList.Add w
            End If
        ElseIf Not isName(i) Then
            catMask = catMask Or ECH_TokenMask(w)
        End If
    Next i
End Sub

' Normaliza (minusculas, sem acentos) e quebra em palavras. O delimitador
' "|" marca pontuacao (virgula, ponto, hifen, aspas, n.o etc.).
Private Function ECH_Tokenize(ByVal text As String, ByRef toks() As String) As Long
    Dim t As String
    Dim ch As String
    Dim cur As String
    Dim code As Long
    Dim i As Long, L As Long
    Dim n As Long
    Dim skipDot As Boolean

    ReDim toks(1 To 64)
    n = 0

    t = " " & NormalizeForComparison(Left$(text, ECH_MAX_CHARS)) & " "

    ' Abreviacoes comuns de logradouro
    t = Replace(t, " av.", " avenida ")
    t = Replace(t, " r. ", " rua ")
    t = Replace(t, " trav.", " travessa ")
    t = Replace(t, " estr.", " estrada ")
    t = Replace(t, " pca.", " praca ")
    t = Replace(t, " pc.", " praca ")
    t = Replace(t, " al. ", " alameda ")

    L = Len(t)
    cur = ""

    For i = 1 To L
        ch = Mid$(t, i, 1)
        code = AscW(ch)

        If (code >= 97 And code <= 122) Or (code >= 48 And code <= 57) Then
            cur = cur & ch
        ElseIf code = 31 Or code = 173 Then
            ' hifen opcional / soft hyphen: ignora
        Else
            ' ponto separador de milhar (1.234): ignora.
            ' If aninhado de proposito: VBA nao faz curto-circuito em And,
            ' e Mid$ com indice invalido geraria erro.
            skipDot = False
            If code = 46 Then
                If i > 1 And i < L Then
                    If ECH_IsDigitCode(AscW(Mid$(t, i - 1, 1))) Then
                        If ECH_IsDigitCode(AscW(Mid$(t, i + 1, 1))) Then skipDot = True
                    End If
                End If
            End If

            If Not skipDot Then
                If Len(cur) > 0 Then
                    ECH_AddTok toks, n, cur
                    cur = ""
                End If
                If Not ECH_IsSpaceCode(code) Then ECH_AddTok toks, n, "|"
            End If
        End If
    Next i

    If Len(cur) > 0 Then ECH_AddTok toks, n, cur

    ECH_Tokenize = n
End Function

Private Sub ECH_AddTok(ByRef toks() As String, ByRef n As Long, ByVal w As String)
    n = n + 1
    If n > UBound(toks) Then ReDim Preserve toks(1 To UBound(toks) * 2)
    toks(n) = w
End Sub

'================================================================================
' VOCABULARIO
'================================================================================

Private Function ECH_IsStreetType(ByVal w As String) As Boolean
    Select Case w
        Case "rua", "avenida", "travessa", "alameda", "estrada", _
             "praca", "rodovia", "viela", "largo", "beco"
            ECH_IsStreetType = True
    End Select
End Function

' Palavras que encerram o nome do logradouro
Private Function ECH_IsStopWord(ByVal w As String) As Boolean
    Select Case w
        Case "n", "no", "na", "nos", "nas", "em", "ao", "aos", "a", "e", _
             "entre", "esquina", "defronte", "frente", "altura", _
             "proximo", "proxima", "proximos", "junto", "ate", _
             "cruzamento", "com", "bairro", "neste", "nesta", _
             "numero", "sob", "sobre"
            ECH_IsStopWord = True
    End Select
End Function

Private Function ECH_IsArticle(ByVal w As String) As Boolean
    Select Case w
        Case "de", "do", "da", "dos", "das"
            ECH_IsArticle = True
    End Select
End Function

Private Function ECH_CatName(ByVal idx As Long) As String
    Select Case idx
        Case 1: ECH_CatName = "pavimentacao / buracos"
        Case 2: ECH_CatName = "iluminacao publica"
        Case 3: ECH_CatName = "arvores / poda"
        Case 4: ECH_CatName = "mato / limpeza"
        Case 5: ECH_CatName = "sinalizacao / estacionamento"
        Case 6: ECH_CatName = "animais peconhentos"
        Case 7: ECH_CatName = "lixo / containers"
        Case 8: ECH_CatName = "agua / esgoto"
        Case 9: ECH_CatName = "seguranca / policiamento"
    End Select
End Function

' Radicais (ja normalizados) que identificam cada assunto.
' Para incluir novos termos, basta acrescentar o radical na lista da categoria.
Private Function ECH_CatStems(ByVal idx As Long) As String
    Select Case idx
        Case 1: ECH_CatStems = "buraco,cratera,asfalt,recap,pavimen,afundam,erosao,cascalh,tapa"
        Case 2: ECH_CatStems = "lampada,ilumina,luminari,poste"
        Case 3: ECH_CatStems = "poda,arvore,extracao,galho"
        Case 4: ECH_CatStems = "rocagem,mato,capina,limpeza,revitaliz,formigueir,cupinzeir"
        Case 5: ECH_CatStems = "estacionamento,vaga,placa,lombada,sinaliza,pintura,faixa,semaforo"
        Case 6: ECH_CatStems = "escorpi,peconhent,dedetiz"
        Case 7: ECH_CatStems = "container,conteiner,lixo,coleta,cacamba,entulho"
        Case 8: ECH_CatStems = "vazament,bueiro,esgoto,galeria"
        Case 9: ECH_CatStems = "ronda,policiam,guarda,vigilan"
    End Select
End Function

' Mascara de bits das categorias a que a palavra pertence
Private Function ECH_TokenMask(ByVal w As String) As Long
    Dim c As Long
    Dim bit As Long
    Dim s As Variant

    ECH_TokenMask = 0
    If Len(w) < 4 Then Exit Function

    ECH_EnsureStems

    bit = 1
    For c = 1 To ECH_CAT_COUNT
        For Each s In echStems(c)
            If Left$(w, Len(s)) = s Then
                ECH_TokenMask = ECH_TokenMask Or bit
                Exit For
            End If
        Next s
        bit = bit * 2
    Next c
End Function

' Preenche o cache de radicais uma unica vez (um Split por categoria)
Private Sub ECH_EnsureStems()
    Dim c As Long
    If echStemsReady Then Exit Sub
    For c = 1 To ECH_CAT_COUNT
        echStems(c) = Split(ECH_CatStems(c), ",")
    Next c
    echStemsReady = True
End Sub

Private Function ECH_MaskToText(ByVal mask As Long) As String
    Dim c As Long
    Dim bit As Long
    Dim res As String

    bit = 1
    For c = 1 To ECH_CAT_COUNT
        If (mask And bit) <> 0 Then
            If Len(res) > 0 Then res = res & " e "
            res = res & """" & ECH_CatName(c) & """"
        End If
        bit = bit * 2
    Next c

    ECH_MaskToText = res
End Function

'================================================================================
' UTILITARIOS
'================================================================================

' Mesmo logradouro? Igual, um contido no outro, ou quase igual (erro de digitacao)
Private Function ECH_SameName(ByVal a As String, ByVal b As String) As Boolean
    Dim m As Long
    Dim tol As Long

    ECH_SameName = False

    If a = b Then
        ECH_SameName = True
        Exit Function
    End If

    If InStr(a, b) > 0 Or InStr(b, a) > 0 Then
        ECH_SameName = True
        Exit Function
    End If

    m = Len(a)
    If Len(b) < m Then m = Len(b)

    If m >= 8 Then
        tol = 2
    ElseIf m >= 5 Then
        tol = 1
    Else
        tol = 0
    End If

    If LevenshteinDistance(a, b) <= tol Then ECH_SameName = True
End Function

Private Function ECH_IsDigits(ByVal w As String) As Boolean
    If Len(w) = 0 Then Exit Function
    ECH_IsDigits = Not (w Like "*[!0-9]*")
End Function

Private Function ECH_IsDigitCode(ByVal code As Long) As Boolean
    ECH_IsDigitCode = (code >= 48 And code <= 57)
End Function

Private Function ECH_IsSpaceCode(ByVal code As Long) As Boolean
    Select Case code
        Case 32, 9, 10, 11, 12, 13, 7, 160
            ECH_IsSpaceCode = True
    End Select
End Function

Private Function ECH_ColContains(col As Collection, ByVal v As String) As Boolean
    Dim i As Long
    For i = 1 To col.count
        If CStr(col(i)) = v Then
            ECH_ColContains = True
            Exit Function
        End If
    Next i
End Function

Private Function ECH_ListStreets(names As Collection, kinds As Collection) As String
    Dim i As Long
    Dim res As String
    For i = 1 To names.count
        If Len(res) > 0 Then res = res & "; "
        res = res & kinds(i) & " " & names(i)
    Next i
    ECH_ListStreets = res
End Function

Private Function ECH_ListPlain(col As Collection) As String
    Dim i As Long
    Dim res As String
    For i = 1 To col.count
        If Len(res) > 0 Then res = res & ", "
        res = res & col(i)
    Next i
    ECH_ListPlain = res
End Function

'================================================================================
' AUTOTESTE (nao precisa de documento aberto)
'================================================================================
' Casos reais das indicacoes 4163, 4226, 4258 e 4301 (devem ser detectados)
' e 4300 e 4218 (devem passar sem aviso).
Public Sub TestarCoerenciaEmenta()
    Dim cabecalho As String
    Dim r As String
    Dim ok As Long, total As Long

    cabecalho = "Nos termos do Art. 108 do Regimento Interno desta Casa de Leis, " & _
                "dirijo-me a Vossa Excelencia para indicar que, por intermedio do setor competente, "

    Dim casos(1 To 6, 1 To 3) As String

    casos(1, 1) = "4163 (deve avisar)"
    casos(1, 2) = "Indica ao Poder Executivo Municipal a possibilidade de instalacao de vagas de " & _
                  "estacionamento destinadas a motocicletas na Rua Parana, em frente a Padaria, " & _
                  "esquina com Rua Ceara, neste municipio."
    casos(1, 3) = cabecalho & "seja realizada operacao de manutencao de afundamento do asfalto na " & _
                  "Rua da Voa Vontade, 138, Jardim Vista Alegre, neste municipio."

    casos(2, 1) = "4226 (deve avisar)"
    casos(2, 2) = "Indica ao Poder Executivo Municipal que realize o servico de tapa-buracos na " & _
                  "Rua do Cromo, no cruzamento com a Rua Jose Jorge Patricio, no bairro Jardim Mollon 4."
    casos(2, 3) = cabecalho & "sejam realizados os servicos de tapa-buracos na Rua Joao Sartori, " & _
                  "n.o 859, no bairro Jd. Mollon 4."

    casos(3, 1) = "4258 (deve avisar)"
    casos(3, 2) = "Sugere ao Poder Executivo Municipal troca de lampada queimada, defronte o " & _
                  "n. 444 da Rua Gentil Pavan, no bairro Vila Rica."
    casos(3, 3) = cabecalho & "troca de lampada queimada, defronte o n. 740 da Rua Joao Pessoa, " & _
                  "no bairro Planalto do Sol."

    casos(4, 1) = "4301 (deve avisar)"
    casos(4, 2) = "Indica ao Poder Executivo Municipal que realize o servico de tapa-buracos na " & _
                  "Rua do Cacau, n. 483, bairro Cidade Nova, onde o asfalto esta afundando."
    casos(4, 3) = cabecalho & "sejam realizados os servicos de tapa-buracos na Rua do Milho, " & _
                  "n. 483, bairro Cidade Nova, visto que a via esta cedendo."

    casos(5, 1) = "4300 (deve passar)"
    casos(5, 2) = "Indica ao Poder Executivo Municipal que realize o servico de tapa-buracos na " & _
                  "Rua do Milho, n. 546, bairro Cidade Nova, onde o asfalto esta afundando."
    casos(5, 3) = cabecalho & "sejam realizados os servicos de tapa-buracos na Rua do Milho, " & _
                  "n. 546, bairro Cidade Nova, visto que o local esta cedendo."

    casos(6, 1) = "4218 (deve passar)"
    casos(6, 2) = "Indica ao Poder Executivo Municipal que realize o servico de tapa-buracos na " & _
                  "Rua do Cromo, n. 1232, no bairro Jd. Mollon 4."
    casos(6, 3) = cabecalho & "sejam realizados os servicos de tapa-buracos na Rua do Cromo, " & _
                  "n. 1232, no bairro Jd. Mollon 4."

    Dim i As Long
    Dim esperaAviso As Boolean
    Dim teveAviso As Boolean
    Dim out As String

    For i = 1 To 6
        esperaAviso = (InStr(casos(i, 1), "deve avisar") > 0)
        r = ECH_CompareTexts(casos(i, 2), casos(i, 3))
        teveAviso = (Len(r) > 0)
        total = total + 1
        If esperaAviso = teveAviso Then
            ok = ok + 1
            out = out & "OK   - " & casos(i, 1) & vbCrLf
        Else
            out = out & "FALHA - " & casos(i, 1) & vbCrLf
        End If
        If teveAviso Then out = out & "        " & Replace(r, vbCrLf, vbCrLf & "        ") & vbCrLf
    Next i

    MsgBox ok & " de " & total & " casos conforme o esperado." & vbCrLf & vbCrLf & out, _
           IIf(ok = total, vbInformation, vbExclamation), ECH_TITLE & " - Autoteste"
End Sub
