#requires -Version 5.1

Import-Module Pester -ErrorAction Stop

. $PSScriptRoot\Helpers.ps1



Describe 'Z7_STDPROPOSERS - Mod_12_AIStructure' {

    BeforeAll {

        $repoRoot = Get-RepoRoot

        $mod12Path = Join-Path $repoRoot 'source\main\Mod_12_AIStructure.bas'

        if (-not (Test-Path $mod12Path)) {

            throw "Arquivo nao encontrado: $mod12Path"

        }



        $script:mod12Content = Get-Content $mod12Path -Raw -Encoding UTF8

    }



    Context 'Estrutura do modulo' {

        It 'Tem Option Explicit' {

            $script:mod12Content | Should Match '(?m)^Option Explicit'

        }



        It 'Tem Attribute VB_Name correto' {

            $script:mod12Content | Should Match 'Attribute VB_Name = "Mod_12_AIStructure"'

        }



        It 'Funcoes e Subs estao balanceados' {

            $functionStarts = ([regex]::Matches($script:mod12Content, '(?m)^(Public |Private )?Function\s+\w+')).Count

            $functionEnds = ([regex]::Matches($script:mod12Content, '(?m)^End Function')).Count

            $subStarts = ([regex]::Matches($script:mod12Content, '(?m)^(Public |Private )?Sub\s+\w+')).Count

            $subEnds = ([regex]::Matches($script:mod12Content, '(?m)^End Sub')).Count



            $functionStarts | Should Be $functionEnds

            $subStarts | Should Be $subEnds

        }

    }



    Context 'Funcoes publicas de entrada' {

        It 'Declara IdentifyDocumentStructureWithAI' {

            $script:mod12Content | Should Match '(?m)^Public Function IdentifyDocumentStructureWithAI\(doc As Document\) As Boolean'

        }



        It 'Declara DiagnosticarEstruturaIA' {

            $script:mod12Content | Should Match '(?m)^Public Sub DiagnosticarEstruturaIA\(\)'

        }



        It 'Declara TestarEstruturaIADocumentoAtual' {

            $script:mod12Content | Should Match '(?m)^Public Sub TestarEstruturaIADocumentoAtual\(\)'

        }

    }



    Context 'Funcoes privadas de infraestrutura' {

        It 'Declara MontarTextoDocumentoParaIA' {

            $script:mod12Content | Should Match '(?m)^Private Function MontarTextoDocumentoParaIA\(doc As Document\) As String'

        }



        It 'Declara MontarPromptEstrutura' {

            $script:mod12Content | Should Match '(?m)^Private Function MontarPromptEstrutura\(\) As String'

        }



        It 'Declara MontarJSONPayload' {

            $script:mod12Content | Should Match '(?m)^Private Function MontarJSONPayload\('

        }



        It 'Declara AI_ChamarAPI' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_ChamarAPI\('

        }



        It 'Declara ParsearRespostaEstruturaIA' {

            $script:mod12Content | Should Match '(?m)^Private Function ParsearRespostaEstruturaIA\('

        }



        It 'Declara ValidarIndicesEstrutura' {

            $script:mod12Content | Should Match '(?m)^Private Function ValidarIndicesEstrutura\(doc As Document\) As Boolean'

        }



        It 'Declara MarcarFlagsEstrutura' {

            $script:mod12Content | Should Match '(?m)^Private Sub MarcarFlagsEstrutura\(doc As Document\)'

        }

    }



    Context 'Funcoes de parse JSON' {

        It 'Declara AI_ExtrairIndiceUnico' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_ExtrairIndiceUnico\('

        }



        It 'Declara AI_ExtrairArrayPrimeiro' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_ExtrairArrayPrimeiro\('

        }



        It 'Declara AI_ExtrairArrayUltimo' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_ExtrairArrayUltimo\('

        }



        It 'Declara AI_ExtrairContentJSON' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_ExtrairContentJSON\('

        }



        It 'Declara AI_DesescaparJSON' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_DesescaparJSON\('

        }



        It 'Declara EscaparJSONAI' {

            $script:mod12Content | Should Match '(?m)^Private Function EscaparJSONAI\('

        }

    }



    Context 'Funcoes de conversao e criptografia' {

        It 'Declara AI_StringParaUTF8' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_StringParaUTF8\('

        }



        It 'Declara AI_BytesParaStringUTF8' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_BytesParaStringUTF8\('

        }



        It 'AI_BytesParaStringUTF8 aceita Variant (compativel com ResponseBody)' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_BytesParaStringUTF8\(.+As Variant'

        }



        It 'Declara AI_CarregarChaveAPI' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_CarregarChaveAPI\(\) As String'

        }



        It 'Declara AI_CarregarModelo' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_CarregarModelo\(\) As String'

        }



        It 'AI_StringParaUTF8 retorna Variant (nao Byte())' {

            $script:mod12Content | Should Match '(?m)^Private Function AI_StringParaUTF8\(.+\) As Variant'

        }

    }



    Context 'Constantes do modulo' {

        It 'Define AI_STRUCT_URL para OpenRouter' {

            $script:mod12Content | Should Match 'AI_STRUCT_URL'

            $script:mod12Content | Should Match 'https://openrouter\.ai/api/v1/chat/completions'

        }



        It 'Define AI_STRUCT_PREFIX' {

            $script:mod12Content | Should Match 'AI_STRUCT_PREFIX.*"AI_STRUCTURE"'

        }



        It 'Define AI_STRUCT_DEFAULT_MODEL' {

            $script:mod12Content | Should Match 'AI_STRUCT_DEFAULT_MODEL'

        }



        It 'Define timeouts HTTP' {

            $script:mod12Content | Should Match 'AI_STRUCT_RESOLVE_TIMEOUT'

            $script:mod12Content | Should Match 'AI_STRUCT_CONNECT_TIMEOUT'

            $script:mod12Content | Should Match 'AI_STRUCT_SEND_TIMEOUT'

            $script:mod12Content | Should Match 'AI_STRUCT_RECEIVE_TIMEOUT'

        }



        It 'Define limites de protecao de contexto' {

            $script:mod12Content | Should Match 'MAX_PARAGRAPHS_FOR_AI'

            $script:mod12Content | Should Match 'MAX_PARAGRAPH_TEXT_LENGTH'

        }



        It 'Define LOG_LEVEL_DEBUG' {

            $script:mod12Content | Should Match 'LOG_LEVEL_DEBUG'

        }

    }



    Context 'Declaracoes Windows API (DPAPI)' {

        It 'Declara tipo AI_DATA_BLOB' {

            $script:mod12Content | Should Match 'Private Type AI_DATA_BLOB'

        }



        It 'Declara AI_CryptUnprotectData com PtrSafe para VBA7' {

            $script:mod12Content | Should Match 'Declare PtrSafe Function AI_CryptUnprotectData'

        }



        It 'Declara AI_CopyMemory com PtrSafe para VBA7' {

            $script:mod12Content | Should Match 'Declare PtrSafe Sub AI_CopyMemory'

        }



        It 'Declara AI_LocalFree com PtrSafe para VBA7' {

            $script:mod12Content | Should Match 'Declare PtrSafe Function AI_LocalFree'

        }



        It 'Tem fallback para VBA6 (sem PtrSafe)' {

            $script:mod12Content | Should Match '#Else'

            $script:mod12Content | Should Match 'Declare Function AI_CryptUnprotectData'

        }

    }



    Context 'Logging e observabilidade' {

        It 'Usa LogMessage do projeto' {

            $script:mod12Content | Should Match 'LogMessage'

        }



        It 'Usa LogSection para secoes de log' {

            $script:mod12Content | Should Match 'LogSection'

        }



        It 'Usa LogStepStart para inicio de etapas' {

            $script:mod12Content | Should Match 'LogStepStart'

        }



        It 'Usa LogStepComplete para conclusao de etapas' {

            $script:mod12Content | Should Match 'LogStepComplete'

        }



        It 'Usa LogStepSkipped para etapas ignoradas' {

            $script:mod12Content | Should Match 'LogStepSkipped'

        }



        It 'Usa LogMetric para metricas' {

            $script:mod12Content | Should Match 'LogMetric'

        }



        It 'Usa LOG_LEVEL_INFO' {

            $script:mod12Content | Should Match 'LOG_LEVEL_INFO'

        }



        It 'Usa LOG_LEVEL_WARNING' {

            $script:mod12Content | Should Match 'LOG_LEVEL_WARNING'

        }



        It 'Usa LOG_LEVEL_ERROR' {

            $script:mod12Content | Should Match 'LOG_LEVEL_ERROR'

        }



        It 'Usa LOG_LEVEL_DEBUG para depuracao detalhada' {

            $script:mod12Content | Should Match 'LOG_LEVEL_DEBUG'

        }



        It 'Registra tempo de execucao HTTP' {

            $script:mod12Content | Should Match 'httpElapsed'

        }



        It 'Registra indices extraidos no parse' {

            $script:mod12Content | Should Match 'Indices extraidos'

        }



        It 'Registra tamanho do payload' {

            $script:mod12Content | Should Match 'Payload montado'

        }



        It 'Registra tamanho do texto montado' {

            $script:mod12Content | Should Match 'Texto montado'

        }



        It 'Registra content extraido da resposta' {

            $script:mod12Content | Should Match 'Content extraido'

        }



        It 'Registra flags marcados no cache' {

            $script:mod12Content | Should Match 'Flags marcados'

        }



        It 'Registra chave API descriptografada' {

            $script:mod12Content | Should Match 'Chave API descriptografada'

        }



        It 'Registra modelo carregado' {

            $script:mod12Content | Should Match 'Modelo carregado do arquivo'

        }



        It 'Registra modelo padrao' {

            $script:mod12Content | Should Match 'Modelo padrao'

        }



        It 'Registra prompt de estrutura montado' {

            $script:mod12Content | Should Match 'Prompt de estrutura montado'

        }



        It 'Registra JSON escapado' {

            $script:mod12Content | Should Match 'JSON escapado'

        }



        It 'Registra validacao de titulo invalido' {

            $script:mod12Content | Should Match 'titulo invalido'

        }



        It 'Registra validacao de ementa fora do range' {

            $script:mod12Content | Should Match 'ementa fora do range'

        }



        It 'Registra HTTP status com tempo' {

            $script:mod12Content | Should Match 'HTTP 200 OK em'

        }



        It 'Registra erro HTTP com numero e descricao' {

            $script:mod12Content | Should Match 'Err\.Number & " - " & Err\.Description'

        }

    }



    Context 'Tratamento de erro' {

        It 'IdentifyDocumentStructureWithAI tem ErrorHandler' {

            $script:mod12Content | Should Match '(?s)Public Function IdentifyDocumentStructureWithAI.*?ErrorHandler:'

        }



        It 'AI_ChamarAPI tem ErrorHandler' {

            $script:mod12Content | Should Match '(?s)Private Function AI_ChamarAPI.*?ErrorHandler:'

        }



        It 'ParsearRespostaEstruturaIA tem ErrorHandler' {

            $script:mod12Content | Should Match '(?s)Private Function ParsearRespostaEstruturaIA.*?ErrorHandler:'

        }



        It 'ValidarIndicesEstrutura tem ErrorHandler' {

            $script:mod12Content | Should Match '(?s)Private Function ValidarIndicesEstrutura.*?ErrorHandler:'

        }



        It 'DiagnosticarEstruturaIA tem ErrorHandler' {

            $script:mod12Content | Should Match '(?s)Public Sub DiagnosticarEstruturaIA.*?ErrorHandler:'

        }



        It 'TestarEstruturaIADocumentoAtual tem ErrorHandler' {

            $script:mod12Content | Should Match '(?s)Public Sub TestarEstruturaIADocumentoAtual.*?ErrorHandler:'

        }



        It 'MontarTextoDocumentoParaIA tem ErrorHandler' {

            $script:mod12Content | Should Match '(?s)Private Function MontarTextoDocumentoParaIA.*?ErrorHandler:'

        }



        It 'AI_CarregarChaveAPI tem ErrorHandler' {

            $script:mod12Content | Should Match '(?s)Private Function AI_CarregarChaveAPI.*?ErrorHandler:'

        }



        It 'ErrorHandler de IdentifyDocumentStructureWithAI loga Err.Number' {

            $script:mod12Content | Should Match 'Erro inesperado.*Err\.Number'

        }



        It 'ErrorHandler de AI_ChamarAPI loga Err.Number' {

            $script:mod12Content | Should Match 'Erro HTTP.*Err\.Number'

        }

    }



    Context 'Integracao com infraestrutura do projeto' {

        It 'Usa GetZ7StdProposersDataPath de Mod_01_Infrastructure' {

            $script:mod12Content | Should Match 'GetZ7StdProposersDataPath'

        }



        It 'Usa variaveis globais de indices estruturais' {

            $script:mod12Content | Should Match 'tituloParaIndex'

            $script:mod12Content | Should Match 'ementaParaIndex'

            $script:mod12Content | Should Match 'vocativoStartIndex'

            $script:mod12Content | Should Match 'corpoStartIndex'

            $script:mod12Content | Should Match 'tituloJustificativaIndex'

            $script:mod12Content | Should Match 'justificativaStartIndex'

            $script:mod12Content | Should Match 'dataParaIndex'

            $script:mod12Content | Should Match 'assinaturaStartIndex'

            $script:mod12Content | Should Match 'tituloAnexoIndex'

            $script:mod12Content | Should Match 'anexoStartIndex'

        }



        It 'Usa paragraphCache para marcar flags' {

            $script:mod12Content | Should Match 'paragraphCache'

            $script:mod12Content | Should Match 'cacheSize'

        }



        It 'Usa campos do paragraphCache para estrutura' {

            $script:mod12Content | Should Match '\.isTitulo'

            $script:mod12Content | Should Match '\.isEmenta'

            $script:mod12Content | Should Match '\.isVocativo'

            $script:mod12Content | Should Match '\.isCorpoContent'

            $script:mod12Content | Should Match '\.isTituloJustificativa'

            $script:mod12Content | Should Match '\.isJustificativaContent'

            $script:mod12Content | Should Match '\.isData'

            $script:mod12Content | Should Match '\.isAssinatura'

            $script:mod12Content | Should Match '\.isTituloAnexo'

            $script:mod12Content | Should Match '\.isAnexoContent'

        }



        It 'Usa MSXML2.ServerXMLHTTP.6.0 para HTTP' {

            $script:mod12Content | Should Match 'MSXML2\.ServerXMLHTTP\.6\.0'

        }



        It 'Usa ADODB.Stream para conversao UTF-8' {

            $script:mod12Content | Should Match 'ADODB\.Stream'

        }



        It 'Usa VBScript.RegExp para parse JSON' {

            $script:mod12Content | Should Match 'VBScript\.RegExp'

        }

    }



    Context 'Prompt de estrutura' {

        It 'Prompt contem instrucao para retornar JSON' {

            $script:mod12Content | Should Match 'Retorne APENAS o JSON'

        }



        It 'Prompt contem campos obrigatorios do JSON' {

            $script:mod12Content | Should Match '"titulo"'

            $script:mod12Content | Should Match '"ementa"'

            $script:mod12Content | Should Match '"vocativo"'

            $script:mod12Content | Should Match '"corpo"'

            $script:mod12Content | Should Match '"titulo_da_justificativa"'

            $script:mod12Content | Should Match '"justificativa"'

            $script:mod12Content | Should Match '"data"'

            $script:mod12Content | Should Match '"assinatura"'

            $script:mod12Content | Should Match '"titulo_do_anexo"'

            $script:mod12Content | Should Match '"anexo"'

        }



        It 'Prompt contem regras de identificacao' {

            $script:mod12Content | Should Match 'REGRAS:'

            $script:mod12Content | Should Match 'Valores sao NUMEROS dos paragrafos'

        }

    }



    Context 'Diagnostico e teste' {

        It 'DiagnosticarEstruturaIA testa conectividade HTTP' {

            $script:mod12Content | Should Match '(?s)DiagnosticarEstruturaIA.*?openrouter\.ai/api/v1/models'

        }



        It 'DiagnosticarEstruturaIA exibe MsgBox com resultado' {

            $script:mod12Content | Should Match '(?s)DiagnosticarEstruturaIA.*?MsgBox'

        }



        It 'DiagnosticarEstruturaIA usa LogSection' {

            $script:mod12Content | Should Match '(?s)DiagnosticarEstruturaIA.*?LogSection.*?DIAGNOSTICO ESTRUTURA IA'

        }



        It 'TestarEstruturaIADocumentoAtual reseta indices antes do teste' {

            $script:mod12Content | Should Match '(?s)TestarEstruturaIADocumentoAtual.*?tituloParaIndex = 0'

        }



        It 'TestarEstruturaIADocumentoAtual chama IdentifyDocumentStructureWithAI' {

            $script:mod12Content | Should Match '(?s)TestarEstruturaIADocumentoAtual.*?IdentifyDocumentStructureWithAI\(doc\)'

        }



        It 'TestarEstruturaIADocumentoAtual exibe resultados em MsgBox' {

            $script:mod12Content | Should Match '(?s)TestarEstruturaIADocumentoAtual.*?Estrutura identificada com sucesso'

        }



        It 'TestarEstruturaIADocumentoAtual exibe tempo de execucao' {

            $script:mod12Content | Should Match '(?s)TestarEstruturaIADocumentoAtual.*?Format\(elapsed'

        }

    }



    Context 'Validacao de indices de estrutura' {

        It 'Valida titulo (obrigatorio)' {

            $script:mod12Content | Should Match 'tituloParaIndex <= 0 Or tituloParaIndex > maxPara'

        }



        It 'Valida ementa' {

            $script:mod12Content | Should Match 'ementaParaIndex > maxPara'

        }



        It 'Valida vocativo (start <= end)' {

            $script:mod12Content | Should Match 'vocativoStartIndex > vocativoEndIndex'

        }



        It 'Valida corpo (start <= end)' {

            $script:mod12Content | Should Match 'corpoStartIndex > corpoEndIndex'

        }



        It 'Valida justificativa (start <= end)' {

            $script:mod12Content | Should Match 'justificativaStartIndex > justificativaEndIndex'

        }



        It 'Valida assinatura (start <= end)' {

            $script:mod12Content | Should Match 'assinaturaStartIndex > assinaturaEndIndex'

        }



        It 'Valida anexo (start <= end)' {

            $script:mod12Content | Should Match 'anexoStartIndex > anexoEndIndex'

        }

    }

}