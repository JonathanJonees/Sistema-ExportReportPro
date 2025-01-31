<#
=============================================================================================
Nome:           Obter informações sobre os métodos de autenticação multifator (MFA) dos usuários no Microsoft Entra ID
Descrição:      Este script instala automaticamente todos os módulos necessários (mediante sua confirmação) e conecta aos serviços.
                Além disso, ele se conecta ao Microsoft Graph API para obter informações de usuários e exporta os dados de MFA para um arquivo CSV.
				
				1. Conectar à Microsoft Graph API: Utiliza Connect-MgGraph com as permissões necessárias.
				2. Criar uma variável com a data: Usa Get-Date para formatar a data.
				3. Definir o local do arquivo CSV: Especifica o caminho onde o relatório será salvo.
				4. Obter usuários: Usa Get-MgBetaUser para obter os primeiros 50 usuários.
				5. Inicializar variáveis: Cria uma lista para armazenar os resultados e define contadores para o progresso.
				6. Iterar sobre os usuários: Para cada usuário, calcula o progresso e cria um objeto para armazenar as informações de MFA.
				7. Checar métodos de autenticação: Verifica os métodos de autenticação de cada usuário e atualiza o status de MFA.
				8. Adicionar resultados à lista: Adiciona o objeto de usuário à lista de resultados.
				9. Exportar para CSV: Salva os resultados em um arquivo CSV.

Versão:         1.0
Autor:          Iago Rufino
=============================================================================================
#>
# Conectar ao Microsoft Graph API
Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All"
# Criar uma variavel com a data
$LogDate = Get-Date -f yyyyMMddhhmm
# Definir o local da variavel CVS
$Csvfile = "C:\Sistema-ExportReportPro\Relatorios\Mfa\MFAUsers_$LogDate.csv"
# Get em todos os usuarios Microsoft Entra ID usando o Microsoft Graph Beta API
$users = Get-MgBetaUser -Top 50
# Inicia o aro
$results = @()
# Inicia a contagem de progresso
$counter = 0
$totalUsers = $users.Count
# foreach em todos os usuarios
foreach ($user in $users) {
    $counter++
    # Calcula a porcentagem de progresso
    $percentComplete = [math]::Round(($counter / $totalUsers) * 100)
    # Definir a barra de progresso com a UPN
    $progressParams = @{
        Activity        = "Processing Users"
        Status          = "User $($counter) of $totalUsers - $($user.UserPrincipalName) - $percentComplete% Complete"
        PercentComplete = $percentComplete
    }
    Write-Progress @progressParams
    # Criar um objeto para salvar as informacoes de MFA
    $myObject = [PSCustomObject]@{
        DisplayName               = "-"
        UserPrincipalName         = "-"
        MFAstatus                 = "Disabled"  # Initialize to "Disabled"
        Email                     = "-"
        Fido2                     = "-"
        MicrosoftAuthenticatorApp = "-"
        Phone                     = "-"
        PhoneNumber               = "-"
        PhoneType               = "-"
        SoftwareOath              = "-"
        TemporaryAccessPass       = "-"
        WindowsHelloForBusiness   = "-"
    }
    $myObject.UserPrincipalName = $user.UserPrincipalName
    $myObject.DisplayName = $user.DisplayName
    $myObject.PhoneNumber = Get-MgUserAuthenticationPhoneMethod -UserId $user.UserPrincipalName | Select -ExpandProperty PhoneNumber
    $myObject.PhoneType = Get-MgUserAuthenticationPhoneMethod -UserId $user.UserPrincipalName | Select -ExpandProperty PhoneType
    # Checar o metodo de autorizacao para cada usuario
    $MFAData = Get-MgBetaUserAuthenticationMethod -UserId $user.Id
    foreach ($method in $MFAData) {
        Switch ($method.AdditionalProperties["@odata.type"]) {
            "#microsoft.graph.emailAuthenticationMethod" {
                $myObject.Email = $true
                $myObject.MFAstatus = "Enabled"
            }
            "#microsoft.graph.fido2AuthenticationMethod" {
                $myObject.Fido2 = $true
                $myObject.MFAstatus = "Enabled"
            }
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                $myObject.MicrosoftAuthenticatorApp = $true
                $myObject.MFAstatus = "Enabled"
            }
            "#microsoft.graph.phoneAuthenticationMethod" {
                $myObject.Phone = $true
                $myObject.MFAstatus = "Enabled"
            }
            "#microsoft.graph.softwareOathAuthenticationMethod" {
                $myObject.SoftwareOath = $true
                $myObject.MFAstatus = "Enabled"
            }
            "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                $myObject.TemporaryAccessPass = $true
                $myObject.MFAstatus = "Enabled"
            }
            "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                $myObject.WindowsHelloForBusiness = $true
                $myObject.MFAstatus = "Enabled"
            }
        }
    }
    # Adicionar o usuario para o aro de resultados
    $results += $myObject
}
# Limpar a barra de progresso
Write-Progress -Activity "Processando usuarios" -Completed
# Exportar os resultados para um CSV
$results | Export-Csv -Path $Csvfile -NoTypeInformation -Encoding UTF8

Clear-Host
   Write-Output "Exportação concluída com sucesso!"
   Write-Host "Valide o Relatório Exportado --> Sistema-ExportReportPro\Relatorios\"  -ForegroundColor Green 
  #Write-Host "Script completo, resultados exportados para $Csvfile." -ForegroundColor Cyan

  # Animação espera
  animacao_espera
  # Menu principal
  Show-Menu