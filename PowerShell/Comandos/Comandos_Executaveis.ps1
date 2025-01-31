<#
=============================================================================================
Nome:           
Descrição:      

Versão:         1.0

1. Define as OUs de onde os computadores serão exportados.
2. Inicializa uma lista para armazenar todos os computadores.


Autor:          Jonathan Jones
=============================================================================================
#>
#Import-Module C:\TEMP\Automacao\Sistema_PoweShell\ExportReportPro.ps1

# Exibe a versão do Script
$Versao = "V1.0.0"  
Clear-Host
Write-Host "==========================================================================================================" -ForegroundColor Cyan
Write-Host "                            ESTA OPÇÃO ESTÁ EM CONSTRUÇÃO                                   $Versao       "                                         
Write-Host 
Write-Host "                  AGUARDE VOCÊ SERA REDIRECIONADO PARA O MENU INICIAL											  "                                         
Write-Host "==========================================================================================================" -ForegroundColor Cyan

# Animação espera
animacao_espera
Start-Sleep -Seconds 10

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
# Menu principal
Show-Menu