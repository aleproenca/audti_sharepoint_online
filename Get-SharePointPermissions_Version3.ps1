<#
.SYNOPSIS
    Lista as permissões de um site do SharePoint, bibliotecas e todas as subpastas.
    Compatível com PowerShell 5.1 (módulo SharePointPnPPowerShellOnline).
    Exporta os resultados para um ficheiro Excel (.xlsx) com múltiplas abas.

.PARAMETER SiteUrl
    URL do site do SharePoint. Ex: https://contoso.sharepoint.com/sites/meusite

.PARAMETER OutputPath
    Caminho do ficheiro Excel de saída.

.PARAMETER IncluirPermissoesHerdadas
    Se ativado, inclui também pastas que herdam permissões do nível superior.

.EXAMPLE
    .\Get-SharePointPermissions.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/meusite"

.EXAMPLE
    .\Get-SharePointPermissions.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/meusite" -IncluirPermissoesHerdadas
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "PermissoesSharePoint_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",

    [Parameter(Mandatory = $false)]
    [switch]$IncluirPermissoesHerdadas
)

# ============================================================
# Verificar módulos necessários
# ============================================================
$modulosNecessarios = @("SharePointPnPPowerShellOnline", "ImportExcel")
foreach ($modulo in $modulosNecessarios) {
    if (-not (Get-Module -Name $modulo -ListAvailable)) {
        Write-Host "ERRO: Módulo '$modulo' não encontrado." -ForegroundColor Red
        Write-Host "Instale com: Install-Module -Name $modulo -Scope CurrentUser -Force" -ForegroundColor Yellow
        exit 1
    }
}

Import-Module SharePointPnPPowerShellOnline -DisableNameChecking
Import-Module ImportExcel

# ============================================================
# Função para obter permissões de um objeto (site, lista, pasta)
# ============================================================
function Get-ObjectPermissions {
    param(
        [Parameter(Mandatory = $true)]
        $ClientObject,

        [Parameter(Mandatory = $true)]
        [string]$Escopo,

        [Parameter(Mandatory = $true)]
        [string]$Recurso,

        [Parameter(Mandatory = $false)]
        [string]$CaminhoCompleto = ""
    )

    $permissions = @()

    try {
        $roleAssignments = Get-PnPProperty -ClientObject $ClientObject -Property RoleAssignments -ErrorAction Stop
        $hasUnique = Get-PnPProperty -ClientObject $ClientObject -Property HasUniqueRoleAssignments -ErrorAction Stop

        foreach ($ra in $roleAssignments) {
            Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings | Out-Null

            $roleNames = ($ra.RoleDefinitionBindings | ForEach-Object { $_.Name }) -join "; "

            # Ignorar "Limited Access" pois é atribuído automaticamente
            if ($roleNames -eq "Limited Access") { continue }

            $permissions += [PSCustomObject]@{
                Escopo           = $Escopo
                Recurso          = $Recurso
                CaminhoCompleto  = $CaminhoCompleto
                PermissoesUnicas = $hasUnique
                TipoMembro       = [string]$ra.Member.PrincipalType
                Membro           = $ra.Member.Title
                LoginMembro      = $ra.Member.LoginName
                NiveisPermissao  = $roleNames
            }
        }
    }
    catch {
        Write-Host "         [AVISO] Não foi possível ler permissões de: $CaminhoCompleto - $($_.Exception.Message)" -ForegroundColor DarkYellow
    }

    return $permissions
}

# ============================================================
# Função recursiva para percorrer todas as subpastas
# ============================================================
function Get-FolderPermissionsRecursive {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LibraryTitle,

        [Parameter(Mandatory = $true)]
        [string]$FolderRelativeUrl,

        [Parameter(Mandatory = $true)]
        [int]$Depth,

        [Parameter(Mandatory = $false)]
        [int]$MaxDepth = 50
    )

    $allFolderPermissions = @()

    if ($Depth -gt $MaxDepth) {
        Write-Host "         [AVISO] Profundidade máxima atingida em: $FolderRelativeUrl" -ForegroundColor DarkYellow
        return $allFolderPermissions
    }

    try {
        # Obter subpastas
        $subFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderRelativeUrl -ItemType Folder -ErrorAction Stop

        foreach ($folder in $subFolders) {
            # Ignorar pastas de sistema
            if ($folder.Name -eq "Forms" -or $folder.Name -like "_*") { continue }

            $folderRelPath = "$FolderRelativeUrl/$($folder.Name)"
            $indent = "         " + ("  " * $Depth)
            Write-Host "$indent -> $folderRelPath" -ForegroundColor DarkGray

            try {
                # Obter o ListItem associado à pasta para verificar permissões
                $folderItem = Get-PnPFolder -Url $folderRelPath -Includes ListItemAllFields.HasUniqueRoleAssignments, ListItemAllFields.RoleAssignments -ErrorAction Stop

                if ($null -ne $folderItem.ListItemAllFields) {
                    $hasUnique = $folderItem.ListItemAllFields.HasUniqueRoleAssignments

                    # Se a pasta tem permissões únicas OU se queremos ver todas
                    if ($hasUnique -or $IncluirPermissoesHerdadas) {
                        $folderPerms = Get-ObjectPermissions `
                            -ClientObject $folderItem.ListItemAllFields `
                            -Escopo "Pasta" `
                            -Recurso $folder.Name `
                            -CaminhoCompleto $folderRelPath

                        $allFolderPermissions += $folderPerms

                        if ($hasUnique) {
                            Write-Host "$indent    [Permissões únicas encontradas]" -ForegroundColor Magenta
                        }
                    }
                }
            }
            catch {
                Write-Host "$indent    [AVISO] Erro ao ler pasta: $($_.Exception.Message)" -ForegroundColor DarkYellow
            }

            # Recursão para subpastas
            $subPerms = Get-FolderPermissionsRecursive `
                -LibraryTitle $LibraryTitle `
                -FolderRelativeUrl $folderRelPath `
                -Depth ($Depth + 1) `
                -MaxDepth $MaxDepth

            $allFolderPermissions += $subPerms
        }
    }
    catch {
        # Pasta vazia ou sem acesso
        if ($_.Exception.Message -notlike "*does not exist*") {
            Write-Host "         [AVISO] Erro ao listar subpastas de '$FolderRelativeUrl': $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }

    return $allFolderPermissions
}

# ============================================================
# Script principal
# ============================================================
try {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host " Listagem de Permissões do SharePoint"       -ForegroundColor Cyan
    Write-Host " (Incluindo subpastas + Export Excel)"       -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""

    # ----------------------------------------------------------
    # 1. Conectar ao site do SharePoint
    # ----------------------------------------------------------
    Write-Host "[1/5] A conectar ao site: $SiteUrl ..." -ForegroundColor Yellow
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
    Write-Host "       Conexão estabelecida com sucesso!" -ForegroundColor Green
    Write-Host ""

    $sitePermissions = @()
    $libraryPermissions = @()
    $folderPermissions = @()

    # ----------------------------------------------------------
    # 2. Permissões ao nível do site
    # ----------------------------------------------------------
    Write-Host "[2/5] A obter permissões do site..." -ForegroundColor Yellow

    $web = Get-PnPWeb -Includes RoleAssignments, HasUniqueRoleAssignments, Title, Url

    $sitePermissions = Get-ObjectPermissions `
        -ClientObject $web `
        -Escopo "Site" `
        -Recurso $web.Title `
        -CaminhoCompleto $web.Url

    Write-Host "       Site: $($web.Title) - $($sitePermissions.Count) atribuições." -ForegroundColor Green
    Write-Host ""

    # ----------------------------------------------------------
    # 3. Permissões das bibliotecas
    # ----------------------------------------------------------
    Write-Host "[3/5] A obter permissões das bibliotecas..." -ForegroundColor Yellow

    $lists = Get-PnPList -Includes BaseType, Hidden, Title, HasUniqueRoleAssignments, RoleAssignments, RootFolder

    $docLibraries = $lists | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }

    Write-Host "       Encontradas $($docLibraries.Count) bibliotecas de documentos." -ForegroundColor Gray

    foreach ($library in $docLibraries) {
        $herancaInfo = if ($library.HasUniqueRoleAssignments) { "Permissões únicas" } else { "Herdadas" }
        Write-Host "       -> $($library.Title) ($herancaInfo)" -ForegroundColor Gray

        $rootFolder = Get-PnPProperty -ClientObject $library -Property RootFolder
        $rootFolderUrl = $rootFolder.ServerRelativeUrl

        $libPerms = Get-ObjectPermissions `
            -ClientObject $library `
            -Escopo "Biblioteca" `
            -Recurso $library.Title `
            -CaminhoCompleto $rootFolderUrl

        $libraryPermissions += $libPerms
    }

    Write-Host "       Total: $($libraryPermissions.Count) atribuições em bibliotecas." -ForegroundColor Green
    Write-Host ""

    # ----------------------------------------------------------
    # 4. Permissões das subpastas (recursivo)
    # ----------------------------------------------------------
    Write-Host "[4/5] A percorrer subpastas de cada biblioteca..." -ForegroundColor Yellow
    if ($IncluirPermissoesHerdadas) {
        Write-Host "       (Incluindo permissões herdadas)" -ForegroundColor Gray
    }
    else {
        Write-Host "       (Apenas pastas com permissões únicas)" -ForegroundColor Gray
    }
    Write-Host ""

    foreach ($library in $docLibraries) {
        $rootFolder = Get-PnPProperty -ClientObject $library -Property RootFolder
        $rootFolderServerRelUrl = $rootFolder.ServerRelativeUrl

        # Converter ServerRelativeUrl para SiteRelativeUrl
        $webServerRelUrl = (Get-PnPWeb).ServerRelativeUrl
        if ($webServerRelUrl -eq "/") {
            $folderSiteRelUrl = $rootFolderServerRelUrl.TrimStart("/")
        }
        else {
            $folderSiteRelUrl = $rootFolderServerRelUrl.Replace($webServerRelUrl, "").TrimStart("/")
        }

        Write-Host "       Biblioteca: $($library.Title) ($folderSiteRelUrl)" -ForegroundColor Cyan

        $fPerms = Get-FolderPermissionsRecursive `
            -LibraryTitle $library.Title `
            -FolderRelativeUrl $folderSiteRelUrl `
            -Depth 0

        $folderPermissions += $fPerms

        Write-Host "       Subpastas com permissões: $($fPerms.Count) atribuições encontradas." -ForegroundColor Green
        Write-Host ""
    }

    # ----------------------------------------------------------
    # 5. Exportar resultados para Excel
    # ----------------------------------------------------------
    Write-Host "[5/5] A exportar resultados para Excel..." -ForegroundColor Yellow

    # Remover ficheiro existente para evitar conflitos
    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    # -- Aba 1: Resumo Geral --
    $resumo = @(
        [PSCustomObject]@{ Informacao = "Site"; Valor = $web.Title }
        [PSCustomObject]@{ Informacao = "URL"; Valor = $SiteUrl }
        [PSCustomObject]@{ Informacao = "Data do relatório"; Valor = (Get-Date -Format "dd/MM/yyyy HH:mm:ss") }
        [PSCustomObject]@{ Informacao = "Permissões do site"; Valor = $sitePermissions.Count }
        [PSCustomObject]@{ Informacao = "Permissões de bibliotecas"; Valor = $libraryPermissions.Count }
        [PSCustomObject]@{ Informacao = "Permissões de subpastas"; Valor = $folderPermissions.Count }
        [PSCustomObject]@{ Informacao = "Total de atribuições"; Valor = ($sitePermissions.Count + $libraryPermissions.Count + $folderPermissions.Count) }
        [PSCustomObject]@{ Informacao = "Bibliotecas analisadas"; Valor = $docLibraries.Count }
        [PSCustomObject]@{ Informacao = "Incluir herdadas"; Valor = $IncluirPermissoesHerdadas.ToString() }
    )

    $resumo | Export-Excel -Path $OutputPath `
        -WorksheetName "Resumo" `
        -AutoSize `
        -BoldTopRow `
        -FreezeTopRow `
        -TableStyle Medium6

    # -- Aba 2: Permissões do Site --
    if ($sitePermissions.Count -gt 0) {
        $sitePermissions | Export-Excel -Path $OutputPath `
            -WorksheetName "Site" `
            -AutoSize `
            -BoldTopRow `
            -FreezeTopRow `
            -TableStyle Medium2
    }

    # -- Aba 3: Permissões das Bibliotecas --
    if ($libraryPermissions.Count -gt 0) {
        $libraryPermissions | Export-Excel -Path $OutputPath `
            -WorksheetName "Bibliotecas" `
            -AutoSize `
            -BoldTopRow `
            -FreezeTopRow `
            -TableStyle Medium4
    }

    # -- Aba 4: Permissões das Subpastas --
    if ($folderPermissions.Count -gt 0) {
        $folderPermissions | Export-Excel -Path $OutputPath `
            -WorksheetName "Subpastas" `
            -AutoSize `
            -BoldTopRow `
            -FreezeTopRow `
            -TableStyle Medium5
    }

    # -- Aba 5: Tudo junto (visão consolidada) --
    $allPermissions = @()
    $allPermissions += $sitePermissions
    $allPermissions += $libraryPermissions
    $allPermissions += $folderPermissions

    if ($allPermissions.Count -gt 0) {
        $allPermissions | Export-Excel -Path $OutputPath `
            -WorksheetName "Consolidado" `
            -AutoSize `
            -BoldTopRow `
            -FreezeTopRow `
            -TableStyle Medium9
    }

    $stopwatch.Stop()

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host " Exportação Concluída!" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Ficheiro: $OutputPath" -ForegroundColor White
    Write-Host ""
    Write-Host "  Abas da planilha:" -ForegroundColor White
    Write-Host "    1. Resumo       - Informações gerais do relatório"
    Write-Host "    2. Site         - Permissões ao nível do site ($($sitePermissions.Count))"
    Write-Host "    3. Bibliotecas  - Permissões das bibliotecas ($($libraryPermissions.Count))"
    Write-Host "    4. Subpastas    - Permissões das subpastas ($($folderPermissions.Count))"
    Write-Host "    5. Consolidado  - Todas as permissões juntas ($($allPermissions.Count))"
    Write-Host ""
    Write-Host "  Tempo total: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
    Write-Host ""

    # Abrir o ficheiro Excel automaticamente
    $openFile = Read-Host "Deseja abrir o ficheiro Excel agora? (S/N)"
    if ($openFile -eq "S" -or $openFile -eq "s") {
        Invoke-Item $OutputPath
    }
}
catch {
    Write-Host ""
    Write-Host "ERRO: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack: $($_.ScriptStackTrace)" -ForegroundColor DarkRed
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Write-Host ""
    Write-Host "Desconectado do SharePoint." -ForegroundColor Gray
}