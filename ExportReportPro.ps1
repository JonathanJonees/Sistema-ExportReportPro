<#
=============================================================================================
Nome:           Export Report Pro 2025

Descrição:      Este script é utilizado para exportar relatórios e executar comandos relacionados
				ao gerenciamento de sistemas. Ele apresenta um menu interativo para selecionar 
				diferentes tipos de relatórios e comandos.

Versão:         1.0.0

Autor:          Jonathan Jones
=============================================================================================
#>

# Define a função do menu
function menu-keman {
    Clear-Host
    Write-Host "==================================================================================================================================" -ForegroundColor Cyan
    Write-Ascii -InputObject '   EXPORT REPORT PRO 2025!' -ForegroundColor Yellow                                
    Write-Host "==================================================================================================================================" -ForegroundColor Cyan
    Write-Host " " 
}

# Menu de seleção de script
menu-keman 
Write-Host "                                                                                                                        $Versao " -ForegroundColor Yellow             
Write-Host 
Write-Host "Olá $Env:UserName, Bem-vindo ao Script Export Report Pro !!!"                                                                     -ForegroundColor Yellow
Write-Host 
Write-Host "Um momento, logo o menu será exibido ! "                                                                                          -ForegroundColor Yellow
Write-Host 

# Animação de espera
animacao_espera

Start-Sleep -Seconds 10
clear-host

# Menu principal
Show-Menu
# Exibe a versão do Script
$Versao = "V1.0.0"

# Animação de espera
function animacao_espera {
    $Symbols = [string[]]('||','||','||','||','||','||','||','||','||','||','||','||','||','||','||','||','||','||','||','||','||','||')
    $SymbolIndex = [byte] 0
    $Job = Start-Job -ScriptBlock { Start-Sleep -Seconds 2 }
    while ($Job.'JobStateInfo'.'State' -eq 'Running') {
        if ($SymbolIndex -ge $Symbols.'Count') {$SymbolIndex = [byte] 0}
        Write-Host -NoNewline -Object ("{0}`b" -f $Symbols[$SymbolIndex++])
        Start-Sleep -Milliseconds 100
    }
}      

# Menu de opções relatorios ou comandos
function Show-Menu {
    param (
        [string]$title = 'EXPORT REPORT PRO'
    )
    Clear-Host
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host "                                $title | MENU PRINCIPAL                                   $Versao         "                                          
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host " " 
    Write-Host "1 - EXPORTAR RELATÓRIOS             " -ForegroundColor Green    
    Write-Host "2 - EXECUÇÃO DE COMANDOS            " -ForegroundColor Green   
    Write-Host 
    Write-Host "20 - SAIR                    21 - SOBRE                     22 - INSTALAR MODULOS                          "  -ForegroundColor Yellow
    Write-Host "===========================================================================================================" -ForegroundColor Cyan
    Write-Host

    $choice = Read-Host "Escolha uma opção"
    switch ($choice) {
        1 { Show-SubMenu }
        2 { & "C:\Sistema-ExportReportPro\PowerShell\Comandos\Comandos_Executaveis.ps1" }
        20 { Exit }
        21 { Show-Sobre }
		22 { Show-Modulos }
        default { Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red }
    }
}

# Menu de opções dos tipos de relatório
function Show-SubMenu {
    param (
        [string]$title = 'MENU PRINCIPAL'
    )
    Clear-Host
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host "                                       $title                                              $Versao        "                                          
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host " " 
    Write-Host "1 - RELATÓRIO DE INFORMAÇÕES AD      " -ForegroundColor Green    
    Write-Host "2 - RELATÓRIOS DO EXCHANGE ON-LINE   " -ForegroundColor Green   
    Write-Host "3 - RELATÓRIOS DO MFA                " -ForegroundColor Green 
    Write-Host "4 - RELATÓRIOS DO TEAMS              " -ForegroundColor Green 
    Write-Host "5 - RELATÓRIOS DO SHAREPOINT         " -ForegroundColor Green
    Write-Host 
    Write-Host "20 - VOLTAR                   21 - SOBRE                     22 - MODULOS                                  "  -ForegroundColor Yellow
    Write-Host "===========================================================================================================" -ForegroundColor Cyan
    Write-Host

    $choice = Read-Host "Escolha uma opção"
    switch ($choice) {
        1 { Show-Activedirectory }
        2 { Show-SubMenuExchange }
        3 { Show-Mfa }
        4 { Show-Teams }
        5 { Show-Sharepoint }
        20 { Show-Menu }
        21 { Show-Sobre }
		22 { Show-Modulos }
        default { Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red }
    }
}
# Submenu Activedirectory 1
function Show-Activedirectory {
    Clear-Host
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host "                                RELATÓRIO DE INFORMAÇÕES AD | SUBMENU                                     "                                          
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host " " 
    Write-Host "1 - RELATÓRIOS DE WKS                   " -ForegroundColor Green    
    Write-Host "2 - RELATÓRIOS DE SERVIDORES            " -ForegroundColor Green   
    Write-Host "3 - RELATÓRIOS DE GRUPOS DISTRIBUIÇÃO   " -ForegroundColor Green 
    Write-Host "4 - RELATÓRIOS DE GRUPOS SEGURANÇA      " -ForegroundColor Green 
    Write-Host "5 - RELATÓRIOS DE USUARIOS              " -ForegroundColor Green 
    Write-Host 
    Write-Host "20 - VOLTAR                                                                                                "  -ForegroundColor Yellow
    Write-Host "===========================================================================================================" -ForegroundColor Cyan
    Write-Host

# Função para exibir mensagem com cor e atraso
function Exibir-Mensagem {
    param (
        [string]$mensagem,
        [string]$cor = 'Green'
    )
    Write-Host $mensagem -ForegroundColor $cor
    Start-Sleep -Seconds 7
    Clear-Host
}

 $subChoice = Read-Host "Escolha uma opção"
    switch ($subChoice) {
        1 { 
        try {
            & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\ActiveDirectory\Relatorio_Workstation.ps1"
        } catch {
            Write-Error "Falha na execução do script: $_"
        }
        }
        2 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\ActiveDirectory\Relatorio_Servidores.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        3 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\ActiveDirectory\Relatorio_Grupo_Distribuicao.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        4 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\ActiveDirectoryy\Relatorio_Grupo_Seguranca.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        5 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\ActiveDirectory\Relatorio_Usuarios.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        20 { Show-Menu }
        default { Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red }
    }
}

# Submenu SubMenuExchange opção 2
function Show-SubMenuExchange {
    Clear-Host
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host "                                RELATÓRIOS DO EXCHANGE ON-LINE | SUBMENU                                  "                                          
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host " " 
    Write-Host "1 - RELATÓRIO DE CAIXAS POSTAIS         " -ForegroundColor Green    
    Write-Host "2 - RELATÓRIO DE CAIXAS COMPARTILHADAS  " -ForegroundColor Green   
    Write-Host "3 - RELATÓRIO DE USUÁRIOS               " -ForegroundColor Green 
    Write-Host 
    Write-Host "20 - VOLTAR                                                                                                "  -ForegroundColor Yellow
    Write-Host "===========================================================================================================" -ForegroundColor Cyan
    Write-Host

    $subChoice = Read-Host "Escolha uma opção"
    switch ($subChoice) {
        1 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\ExchangeOnline\Relatorio_CaixasPostais.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        2 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\ExchangeOnline\Caixas_Compartilhadas.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        3 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\ExchangeOnline\Relatorio_Usuarios.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        20 { Show-Menu }
        default { Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red }
    }
}

# Submenu Show-Mfa opção 3
function Show-Mfa {
    Clear-Host
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host "                                RELATÓRIOS DO MFA | SUBMENU                                  "                                          
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host " " 
    Write-Host "1 - RELATÓRIO DE TODAS AS CONTAS         " -ForegroundColor Green    
    Write-Host "2 - RELATÓRIO DE MFA DESABILITADO        " -ForegroundColor Green   
    Write-Host 
    Write-Host "20 - VOLTAR                                                                                                "  -ForegroundColor Yellow
    Write-Host "===========================================================================================================" -ForegroundColor Cyan
    Write-Host

    $subChoice = Read-Host "Escolha uma opção"
    switch ($subChoice) {
        1 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\Mfa\Contas_Mfa.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        2 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\Mfa\Mfa_Desabilitado.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        20 { Show-Menu }
        default { Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red }
    }
}

# Submenu Show-Teams opção 4
function Show-Teams {
    Clear-Host
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host "                                RELATÓRIOS DO TEAMS | SUBMENU                                  "                                          
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host " " 
    Write-Host "1 - RELATÓRIO DE TODAS AS EQUIPES         " -ForegroundColor Green    
    Write-Host 
    Write-Host "20 - VOLTAR                                                                                                "  -ForegroundColor Yellow
    Write-Host "===========================================================================================================" -ForegroundColor Cyan
    Write-Host

    $subChoice = Read-Host "Escolha uma opção"
    switch ($subChoice) {
        1 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\Teams\Report_todas_as_equipes_ms_teams.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        20 { Show-Menu }
        default { Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red }
    }
}

# Submenu Show-Sharepoint opção 5
function Show-Sharepoint {
    Clear-Host
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host "                                RELATÓRIO DO SHAREPOINT | SUBMENU                                     "                                          
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host " " 
    Write-Host "1 - RELATÓRIOS DE SITES                         " -ForegroundColor Green    
    Write-Host "2 - RELATÓRIOS DE SITES EXTERNOS                " -ForegroundColor Green   
    Write-Host "3 - RELATÓRIOS DE SITES COMPARTILHADOS EXTERNOS " -ForegroundColor Green 
    Write-Host "4 - RELATÓRIOS DE USUARIOS EXTERNOS             " -ForegroundColor Green 
    Write-Host "5 - RELATÓRIOS DE ACESSOS ANONIMOS              " -ForegroundColor Green 
    Write-Host 
    Write-Host "20 - VOLTAR                                                                                                "  -ForegroundColor Yellow
    Write-Host "===========================================================================================================" -ForegroundColor Cyan
    Write-Host

# Função para exibir mensagem com cor e atraso
function Exibir-Mensagem {
    param (
        [string]$mensagem,
        [string]$cor = 'Green'
    )
    Write-Host $mensagem -ForegroundColor $cor
    Start-Sleep -Seconds 10
    Clear-Host
}

 $subChoice = Read-Host "Escolha uma opção"
    switch ($subChoice) {
        1 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\Sharepoint\Relatorio_Sites.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        2 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\Sharepointt\Relatorio_Sites_Externo.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        3 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\Sharepoint\Relatorio_Sites_Internos.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        4 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\Sharepoint\Relatorio_Usuarios_Externos.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        5 { 
            try {
                & "C:\Sistema-ExportReportPro\PowerShell\Relatorios\Sharepoint\Relatorio_Usuarios_Anonimos.ps1"
                Exibir-Mensagem -mensagem "Valide o Relatório Exportado -> Sistema-ExportReportPro\Relatorios\" -cor 'Green'
            } catch {
                Exibir-Mensagem -mensagem "Falha na execução do script: $_" -cor 'Red'
            }
        }
        20 { Show-Menu }
        default { Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red }
    }
}

# Submenu para a opção 22
function Show-Modulos {
    Clear-Host
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host "                                INSTALAR/IMPORTAR MODULOS                              					      "                                          
    Write-Host "==========================================================================================================" -ForegroundColor Cyan

    Write-Host "1 -  EXECUTIONPOLICY UNRESTRICTED       	                 " -ForegroundColor Red    
    Write-Host "2 -  MODULO ACTIVEDIRECTORY             				     " -ForegroundColor Green 
    Write-Host "3 -  MODULO MSONLINE               						     " -ForegroundColor Green  	
    Write-Host "4 -  MODULO MICROSOFT.POWERSHELL.MANAGEMENT                  " -ForegroundColor Green 	
    Write-Host "5 -  MODULO AZURE										     " -ForegroundColor Green 
    Write-Host "6 -  MODULO EXCHANGEONLINEMANAGEMENT                         " -ForegroundColor Green 
    Write-Host "7 -  MODULO MICROSOFT.POWERSHELL.MFA                         " -ForegroundColor Green 
    Write-Host "8 -  MODULO MICROSOFT.ONLINE.SHAREPOINT.POWERSHELL           " -ForegroundColor Green 
    Write-Host "9 -  MODULO MICROSOFTTEAMS                                   " -ForegroundColor Green 
    Write-Host "10 - MODULO MICROSOFT.GRAPH                                  " -ForegroundColor Green 
    Write-Host "11 - MODULO IMPORTEXCEL                                      " -ForegroundColor Green 
    Write-Host "12 - MODULO AZ                                               " -ForegroundColor Green 
    Write-Host "13 - MODULO AZUREADEXPORTER                                  " -ForegroundColor Green 
    Write-Host "14 - MODULO PNP.POWERSHELL                                   " -ForegroundColor Green 
    Write-Host 
    Write-Host "20 - VOLTAR                                                                                                "  -ForegroundColor Yellow
    Write-Host "===========================================================================================================" -ForegroundColor Cyan
    Write-Host
}

# Função para exibir mensagem com cor e atraso
function Exibir-Mensagem {
    param (
        [string]$mensagem,
        [string]$cor = 'Green'
    )
    Write-Host $mensagem -ForegroundColor $cor
    Start-Sleep -Seconds 10
    Clear-Host
}

# Função para importar módulo a partir de um caminho específico
function InstalarOuImportar {
    param (
        [string]$nomeModulo
    )
    $caminhoModulo = "C:\Sistema-ExportReportPro\Modulos\Modulos.ps1"
    if (Test-Path $caminhoModulo) {
        try {
            . $caminhoModulo
            Import-Module $nomeModulo
            Exibir-Mensagem -mensagem "Modulo $nomeModulo importado do caminho $caminhoModulo" -cor 'Green'
        } catch {
            Exibir-Mensagem -mensagem "Falha ao importar o modulo $nomeModulo : $_" -cor 'Red'
        }
    } else {
        Exibir-Mensagem -mensagem "Caminho $caminhoModulo não encontrado" -cor 'Red'
    }
}

$subChoice = Read-Host "Escolha uma opção"
switch ($subChoice) {
    1 { 
        Set-ExecutionPolicy Unrestricted -Scope Process -Force
        Exibir-Mensagem -mensagem "ExecutionPolicy definido para Unrestricted" -cor 'Green'
    }
    2 { InstalarOuImportar-Modulo -nomeModulo "ActiveDirectory" }
    3 { InstalarOuImportar-Modulo -nomeModulo "MSOnline" }
    4 { InstalarOuImportar-Modulo -nomeModulo "Microsoft.PowerShell.Management" }
    5 { InstalarOuImportar-Modulo -nomeModulo "Azure" }
    6 { InstalarOuImportar-Modulo -nomeModulo "ExchangeOnlineManagement" }
    7 { InstalarOuImportar-Modulo -nomeModulo "Microsoft.PowerShell.MFA" }
    8 { InstalarOuImportar-Modulo -nomeModulo "Microsoft.Online.SharePoint.PowerShell" }
    9 { InstalarOuImportar-Modulo -nomeModulo "MicrosoftTeams" }
    10 { InstalarOuImportar-Modulo -nomeModulo "Microsoft.Graph" }
    11 { InstalarOuImportar-Modulo -nomeModulo "ImportExcel" }
    12 { InstalarOuImportar-Modulo -nomeModulo "Az" }
    13 { InstalarOuImportar-Modulo -nomeModulo "AzureADExporter" }
    14 { InstalarOuImportar-Modulo -nomeModulo "PnP.PowerShell" }
    20 { Show-Menu }
    default { Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red }
}


# Submenu para a opção Sobre
function Show-Sobre {
    Clear-Host
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host "                                SCRIPT | EXPORT REPORT PRO 2025" -ForegroundColor Cyan                                                                            
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host " > Nome: Script Export Report Pro"                           -ForegroundColor Gree
    Write-Host " > Nome abreviado: ExportReportPro"                          -ForegroundColor Gree
    Write-Host " > Versão: $Versao"                                          -ForegroundColor Gree
    Write-Host " > Criador: Jonathan Jones"                               	 -ForegroundColor Gree
    Write-Host " > Contato: jonathan.jonees@gmail.com"            			 -ForegroundColor Gree            
    Write-Host 
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
    Write-Host "                                CRÉDITOS | CONTRIBUIDORES" -ForegroundColor Cyan                                                                            
    Write-Host "==========================================================================================================" -ForegroundColor Cyan
	
    Write-Host 

    Write-Host " > Nome: Iago Rufino "                                       -ForegroundColor Gree
    Write-Host " > Contato: iago_tubbos@hotmail.com "                        -ForegroundColor Gree
    Write-Host 
   
    Write-Host "20 - VOLTAR                                                                                                "  -ForegroundColor Yellow
    Write-Host "===========================================================================================================" -ForegroundColor Cyan
    Write-Host

    # Função para exibir mensagem com cor e atraso
function Exibir-Mensagem {
    param (
        [string]$mensagem,
        [string]$cor = 'Green'
    )
    Write-Host $mensagem -ForegroundColor $cor
    Start-Sleep -Seconds 10
    Clear-Host
}

     $choice = Read-Host "Escolha uma opção"
    if ($choice -eq 20) {
        Show-Menu
    } else {
        Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red
        Show-Sobre
    }
}

# Inicia o menu principal
Show-Menu