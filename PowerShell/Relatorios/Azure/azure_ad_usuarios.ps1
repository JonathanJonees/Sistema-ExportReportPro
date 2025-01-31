<#
=============================================================================================
Nome:           Exportar Usuários do Azure AD
Descrição:      Este script instala e importa o módulo Microsoft.Graph, autentica no Microsoft Graph, 
				obtém todos os usuários do Azure AD com todos os atributos e exporta as informações para um arquivo CSV.

Versão:         1.0

1. Instala o módulo Microsoft.Graph.
2. Importa o módulo Microsoft.Graph.
3. Autentica no Microsoft Graph.
4. Obtém todos os usuários do Azure AD com todos os atributos.
5. Exporta os dados dos usuários para um arquivo CSV.

Autor:          Jonathan Jones
=============================================================================================
#>
# Instalar o módulo Microsoft.Graph
Install-Module Microsoft.Graph -Scope CurrentUser

# Importar o módulo Microsoft.Graph
Import-Module Microsoft.Graph

# Autenticar no Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"

# Obter todos os usuários com todos os atributos
$users = Get-MgUser -All

# Diretório e nome do arquivo CSV
$outputDir = "C:\Sistema-ExportReportPro\Relatorios\Azure"
$outputFile = "$outputDir\\azure_ad_users.csv"

# Exportar para CSV
$users | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

Write-Output "Usuários exportados para $outputFile com sucesso."