<#
=============================================================================================
Nome:           Exportar Computadores do Active Directory
Descrição:      Este script obtém todos os computadores de OUs específicas no Active Directory e exporta as informações para um arquivo CSV.

Versão:         1.0

1. Define as OUs de onde os computadores serão exportados.
2. Inicializa uma lista para armazenar todos os computadores.
3. Obtém todos os computadores de cada OU especificada.
4. Exporta os dados dos computadores para um arquivo CSV.

Autor:          Jonathan Jones
=============================================================================================
#>

# Defina as OUs de onde você quer exportar os computadores
$ous = @(
#"OU=ti,OU=DepartmentsS,OU=Corporate,DC=example,DC=local",
#"OU=ti,OU=DepartmentsS,OU=Corporate,DC=example,DC=local"

"OU=Wks Stage Baseline,OU=ORIGINAL CONCESSIONARIAS,OU=GRUPO JSL,DC=jslsa,DC=local",
"OU=Wks UAC Baseline,OU=ORIGINAL CONCESSIONARIAS,OU=GRUPO JSL,DC=jslsa,DC=local"
)

# Inicialize uma lista para armazenar todos os computadores
$allComputers = @()

# Obtenha todos os computadores de cada OU
foreach ($ou in $ous) {
    $computers = Get-ADComputer -Filter * -SearchBase $ou
    $allComputers += $computers
}

# Exporte os computadores para um arquivo CSV
$allComputers | Select-Object Name, DNSHostName, OperatingSystem | Export-Csv -Path "C:\Sistema-ExportReportPro\Relatorios\ActiveDirectory\Relatorio_Workstation.csv" -NoTypeInformation

   Clear-Host
   Write-Output "Exportação concluída com sucesso!"
   Write-Host "Valide o Relatório Exportado --> Sistema-ExportReportPro\Relatorios\"  -ForegroundColor Green 
   
# Animação espera
animacao_espera
# Menu principal
Show-Menu
