<#
=============================================================================================
Nome:           Exportar Usuarios Com Arquivamento Ativo/Inativo
Descrição:      Este script autentica com o Microsoft Graph API usando credenciais de cliente e obtém todos os usuários do Exchange com arquivamento habilitado/desabilitado. 
                Em seguida, filtra os usuários e exporta as informações para um arquivo CSV.

Versão:         1.0

1. Autentica com o Microsoft Graph API usando o fluxo de credenciais de cliente.
2. Obtém todos os usuários do Exchange Online.
3. Filtra os usuários com arquivamento habilitado.
4. Exporta os dados dos usuários filtrados para um arquivo CSV.

Autor:          Jonathan Jones
=============================================================================================
#>

# Autenticação com o Microsoft Graph API
$tenantId = "SEU_TENANT_ID"
$clientId = "SEU_CLIENT_ID"
$clientSecret = "SEU_CLIENT_SECRET"
$tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_id     = $clientId
    client_secret = $clientSecret
}
$response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType "application/x-www-form-urlencoded" -Body $body
$accessToken = $response.access_token

# Obter todos os usuários do Exchange
$uri = "https://graph.microsoft.com/v1.0/users"
$headers = @{
    Authorization = "Bearer $accessToken"
}
$users = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers

# Filtrar usuários com e sem arquivamento habilitado
$archivedUsers = $users.value | Where-Object { $_.archiveStatus -eq "Active" }
$nonArchivedUsers = $users.value | Where-Object { $_.archiveStatus -ne "Active" }

# Definir os caminhos dos arquivos CSV
$csvFilePathArchived = "C:\Sistema-ExportReportPro\Relatorios\ExchangeOnline\usuarios_com_archive.csv"
$csvFilePathNonArchived = "C:\Sistema-ExportReportPro\Relatorios\ExchangeOnline\usuarios_sem_archive.csv"

# Escrever os dados nos arquivos CSV
$archivedUsers | Select-Object displayName, mail, archiveStatus, archiveCreationDate, archiveSize, lastActivityDate, archiveQuota, accountStatus, department, officeLocation | Export-Csv -Path $csvFilePathArchived -NoTypeInformation
$nonArchivedUsers | Select-Object displayName, mail, archiveStatus, archiveCreationDate, archiveSize, lastActivityDate, archiveQuota, accountStatus, department, officeLocation | Export-Csv -Path $csvFilePathNonArchived -NoTypeInformation

   Clear-Host
   Write-Output "Relatórios exportados com sucesso para $csvFilePathArchived e $csvFilePathNonArchived."
   Write-Host "Valide o Relatório Exportado --> Sistema-ExportReportPro\Relatorios\"  -ForegroundColor Green 



# Animação espera
animacao_espera
# Menu principal
Show-Menu
