[CmdletBinding()]
param(
    [Parameter(Mandatory, ParameterSetName = 'Alunos')]
    [switch]$Alunos,

    [Parameter(Mandatory, ParameterSetName = 'Docentes')]
    [switch]$Docentes,

    [Parameter(Mandatory, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Ficheiro
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$timestamp] $Message" -ForegroundColor $Color
}

if (-not (Test-Path -LiteralPath $Ficheiro -PathType Leaf)) {
    throw "O ficheiro CSV '$Ficheiro' nao foi encontrado."
}

$tipoUtilizador = if ($Alunos) { 'alunos' } else { 'docentes' }
$nomeGrupo = if ($Alunos) { 'O365-Alunos' } else { 'O365-Professores' }

try {
    Write-Log 'A importar os modulos Microsoft Graph...' Cyan
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop

    Write-Log 'A ligar ao Microsoft Graph...' Cyan
    Connect-MgGraph -Scopes 'User.ReadWrite.All', 'GroupMember.ReadWrite.All' -NoWelcome -ErrorAction Stop
    Write-Log "Ligacao estabelecida. A criar utilizadores do tipo '$tipoUtilizador' e a adiciona-los ao grupo '$nomeGrupo'." Green

    $nomeGrupoFilter = $nomeGrupo.Replace("'", "''")
    $grupos = @(Get-MgGroup -Filter "displayName eq '$nomeGrupoFilter'" -ErrorAction Stop)
    if ($grupos.Count -eq 0) {
        throw "Nao foi encontrado nenhum grupo com o nome '$nomeGrupo'."
    }

    if ($grupos.Count -gt 1) {
        throw "Foram encontrados $($grupos.Count) grupos com o nome '$nomeGrupo'."
    }

    $grupo = $grupos[0]
    $membrosDoGrupo = @{}
    @(Get-MgGroupMember -GroupId $grupo.Id -All -ErrorAction Stop) | ForEach-Object {
        $membrosDoGrupo[$_.Id] = $true
    }

    $utilizadores = @(Import-Csv -LiteralPath $Ficheiro -Delimiter ';' -ErrorAction Stop)
    if ($utilizadores.Count -eq 0) {
        throw "O ficheiro CSV '$Ficheiro' nao contem registos."
    }

    $processados = 0
    $criados = 0
    $existentes = 0
    $adicionadosAoGrupo = 0
    $jaNoGrupo = 0
    $falhados = 0

    foreach ($utilizador in $utilizadores) {
        $processados++
        $nome = ([string]$utilizador.nome).Trim()
        $apelido = ([string]$utilizador.apelido).Trim()
        $upn = ([string]$utilizador.upn).Trim()
        $password = ([string]$utilizador.password).Trim()

        if ([string]::IsNullOrWhiteSpace($nome) -or
            [string]::IsNullOrWhiteSpace($apelido) -or
            [string]::IsNullOrWhiteSpace($upn) -or
            [string]::IsNullOrWhiteSpace($password)) {
            $falhados++
            Write-Log "Registo $processados ignorado: as colunas 'nome', 'apelido', 'upn' e 'password' sao obrigatorias." Yellow
            continue
        }

        try {
            Write-Log "[$processados/$($utilizadores.Count)] A verificar o utilizador '$upn'..."
            $upnFilter = $upn.Replace("'", "''")
            $existente = @(Get-MgUser -Filter "userPrincipalName eq '$upnFilter'" -ErrorAction Stop)

            if ($existente.Count -gt 0) {
                $existentes++
                $conta = $existente[0]
                Write-Log "O utilizador '$upn' ja existe; a verificar a associacao ao grupo."
            }

            if ($existente.Count -eq 0) {
                $nomeApresentacao = "$nome $apelido"
                $mailNickname = ($upn -split '@')[0]
                $novoUtilizador = @{
                    AccountEnabled = $true
                    DisplayName = $nomeApresentacao
                    GivenName = $nome
                    Surname = $apelido
                    MailNickname = $mailNickname
                    UserPrincipalName = $upn
                    PasswordProfile = @{
                        ForceChangePasswordNextSignIn = $false
                        Password = $password
                    }
                }

                $conta = New-MgUser -BodyParameter $novoUtilizador -ErrorAction Stop
                $criados++
                Write-Log "Criado: '$nomeApresentacao' ($upn)." Green
            }

            if ($membrosDoGrupo.ContainsKey($conta.Id)) {
                $jaNoGrupo++
                Write-Log "O utilizador '$upn' ja pertence ao grupo '$nomeGrupo'." Yellow
                continue
            }

            $referenciaUtilizador = @{
                '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($conta.Id)"
            }
            New-MgGroupMemberByRef -GroupId $grupo.Id -BodyParameter $referenciaUtilizador -ErrorAction Stop
            $membrosDoGrupo[$conta.Id] = $true
            $adicionadosAoGrupo++
            Write-Log "Adicionado: '$upn' ao grupo '$nomeGrupo'." Green
        }
        catch {
            $falhados++
            Write-Log "Falha ao criar '$upn': $($_.Exception.Message)" Red
        }
    }

    Write-Log "Concluido: $processados processados, $criados criados, $existentes ja existiam, $adicionadosAoGrupo adicionados ao grupo, $jaNoGrupo ja pertenciam ao grupo, $falhados falhados." Cyan

    if ($falhados -gt 0) {
        exit 1
    }
}
catch {
    Write-Log "Erro fatal: $($_.Exception.Message)" Red
    exit 1
}