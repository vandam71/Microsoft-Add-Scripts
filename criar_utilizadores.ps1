[CmdletBinding()]
param(
    [Parameter(Mandatory, ParameterSetName = 'Alunos')]
    [switch]$Alunos,

    [Parameter(Mandatory, ParameterSetName = 'Docentes')]
    [switch]$Docentes,

    [Parameter(Mandatory, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Ficheiro,

    [ValidatePattern('^\d{4}-\d{4}$')]
    [string]$AnoLetivo = $(if ((Get-Date).Month -ge 8) { "$(Get-Date -Format 'yyyy')-$((Get-Date).Year + 1)" } else { "$((Get-Date).Year - 1)-$(Get-Date -Format 'yyyy')" })
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

function Ensure-Module {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        $resposta = Read-Host "O modulo '$Name' nao esta instalado. Pretende instala-lo para o utilizador atual? (S/N)"
        if ($resposta -notmatch '^[SsYy]$') {
            throw "O modulo necessario '$Name' nao esta instalado."
        }

        Write-Log "A instalar o modulo '$Name'..." Cyan
        Install-Module -Name $Name -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
    }

    Write-Log "A importar o modulo '$Name'..." Cyan
    Import-Module $Name -ErrorAction Stop
}

function Import-UsersCsv {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $header = Get-Content -LiteralPath $Path -TotalCount 1 -ErrorAction Stop
    $hasSemicolon = $header.Contains(';')
    $hasComma = $header.Contains(',')

    if ($hasSemicolon -eq $hasComma) {
        throw "Nao foi possivel determinar o delimitador do CSV '$Path'. Use apenas ponto e virgula (;) ou virgula (,) no cabecalho."
    }

    $delimiter = if ($hasSemicolon) { ';' } else { ',' }
    Write-Log "CSV detetado com delimitador '$delimiter'." Cyan
    return @(Import-Csv -LiteralPath $Path -Delimiter $delimiter -ErrorAction Stop)
}

function Get-AccountDetails {
    param(
        [Parameter(Mandatory)]
        [psobject]$Record,

        [Parameter(Mandatory)]
        [string]$UserType
    )

    $processo = ([string]$Record.Processo).Trim()
    $nomeCompleto = ([string]$Record.Nome).Trim() -replace '\s+', ' '
    $nif = ([string]$Record.NIF).Trim()

    if ([string]::IsNullOrWhiteSpace($processo) -or
        [string]::IsNullOrWhiteSpace($nomeCompleto) -or
        [string]::IsNullOrWhiteSpace($nif)) {
        throw "As colunas 'Processo', 'Nome' e 'NIF' sao obrigatorias."
    }

    $partesNome = @($nomeCompleto -split ' ' | Where-Object { $_ })
    if ($partesNome.Count -lt 2) {
        throw "O campo 'Nome' deve incluir pelo menos primeiro nome e apelido."
    }

    $nifDigits = $nif -replace '\D', ''
    if ($nifDigits.Length -lt 5) {
        throw "O campo 'NIF' tem de conter pelo menos cinco algarismos."
    }

    $prefixoPassword = if ($UserType -eq 'alunos') { 'Aluno' } else { 'Docente' }
    return [pscustomobject]@{
        Processo = $processo
        Nome = $partesNome[0]
        Apelido = ($partesNome[1..($partesNome.Count - 1)] -join ' ')
        NomeApresentacao = $nomeCompleto
        Upn = "$processo@alunos.amadeo.pt"
        Password = "$prefixoPassword$($nifDigits.Substring(0, 5))#"
    }
}

function Get-TurmaMailNickname {
    param(
        [Parameter(Mandatory)]
        [string]$Ano,

        [Parameter(Mandatory)]
        [string]$Turma,

        [Parameter(Mandatory)]
        [string]$AcademicYear
    )

    $anoNormalizado = $Ano.Trim() -replace '\u00C2\u00BA', '\u00BA'
    $anoMatch = [regex]::Match($anoNormalizado, '\d{1,2}')
    if (-not $anoMatch.Success) {
        throw "O valor '$Ano' na coluna 'Ano' deve incluir pelo menos um numero de ano (por exemplo: '1', '1º ano' ou '1 - turma D')."
    }

    $turmaNormalizada = $Turma.Trim().ToUpperInvariant()
    if ($turmaNormalizada -notmatch '^[A-Z0-9]+$') {
        throw "O valor '$Turma' na coluna 'Turma' deve conter apenas letras ou numeros."
    }

    $anoNumero = $anoMatch.Groups[1].Value
    return "$anoNumero-$anoNumero$turmaNormalizada-$AcademicYear"
}

function Get-OrCreateClassTeam {
    param(
        [Parameter(Mandatory)]
        [string]$MailNickname
    )

    $equipas = @(Get-Team -MailNickName $MailNickname -ErrorAction Stop)
    if ($equipas.Count -gt 1) {
        throw "Foram encontradas $($equipas.Count) equipas com MailNickName '$MailNickname'."
    }

    if ($equipas.Count -eq 1) {
        return [pscustomobject]@{
            Team = $equipas[0]
            Created = $false
        }
    }

    Write-Log "A criar a equipa de turma '$MailNickname' a partir do modelo Turma..." Cyan
    $equipaCriada = New-Team -DisplayName $MailNickname -MailNickname $MailNickname -Template EDU_Class -ErrorAction Stop
    return [pscustomobject]@{
        Team = $equipaCriada
        Created = $true
    }
}

if (-not (Test-Path -LiteralPath $Ficheiro -PathType Leaf)) {
    throw "O ficheiro CSV '$Ficheiro' nao foi encontrado."
}

$tipoUtilizador = if ($Alunos) { 'alunos' } else { 'docentes' }
$nomeGrupo = if ($Alunos) { 'O365-Alunos' } else { 'O365-Professores' }

try {
    Ensure-Module -Name Microsoft.Graph.Users
    Ensure-Module -Name Microsoft.Graph.Groups
    if ($Alunos) {
        Ensure-Module -Name MicrosoftTeams
    }

    Write-Log 'A ligar ao Microsoft Graph...' Cyan
    Connect-MgGraph -Scopes 'User.ReadWrite.All', 'GroupMember.ReadWrite.All' -NoWelcome -ErrorAction Stop
    Write-Log "Ligacao estabelecida. A criar utilizadores do tipo '$tipoUtilizador' e a adiciona-los ao grupo '$nomeGrupo'." Green
    if ($Alunos) {
        Write-Log 'A ligar ao Microsoft Teams para associar os alunos as turmas...' Cyan
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
    }

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

    $utilizadores = @(Import-UsersCsv -Path $Ficheiro)
    if ($utilizadores.Count -eq 0) {
        throw "O ficheiro CSV '$Ficheiro' nao contem registos."
    }

    $processados = 0
    $criados = 0
    $existentes = 0
    $adicionadosAoGrupo = 0
    $jaNoGrupo = 0
    $adicionadosATurma = 0
    $turmasCriadas = 0
    $equipasTurma = @{}
    $falhados = 0

    foreach ($utilizador in $utilizadores) {
        $processados++

        try {
            $contaDetalhes = Get-AccountDetails -Record $utilizador -UserType $tipoUtilizador
            $upn = $contaDetalhes.Upn
            Write-Log "[$processados/$($utilizadores.Count)] A verificar o utilizador '$upn'..."
            $upnFilter = $upn.Replace("'", "''")
            $existente = @(Get-MgUser -Filter "userPrincipalName eq '$upnFilter'" -ErrorAction Stop)

            if ($existente.Count -gt 0) {
                $existentes++
                $conta = $existente[0]
                Write-Log "O utilizador '$upn' ja existe; a verificar a associacao ao grupo."
            }

            if ($existente.Count -eq 0) {
                $mailNickname = ($upn -split '@')[0]
                $novoUtilizador = @{
                    AccountEnabled = $true
                    DisplayName = $contaDetalhes.NomeApresentacao
                    GivenName = $contaDetalhes.Nome
                    Surname = $contaDetalhes.Apelido
                    MailNickname = $mailNickname
                    UserPrincipalName = $upn
                    PasswordProfile = @{
                        ForceChangePasswordNextSignIn = $false
                        Password = $contaDetalhes.Password
                    }
                }

                $conta = New-MgUser -BodyParameter $novoUtilizador -ErrorAction Stop
                $criados++
                Write-Log "Criado: '$($contaDetalhes.NomeApresentacao)' ($upn)." Green
            }

            if ($membrosDoGrupo.ContainsKey($conta.Id)) {
                $jaNoGrupo++
                Write-Log "O utilizador '$upn' ja pertence ao grupo '$nomeGrupo'." Yellow
            }
            else {
                $referenciaUtilizador = @{
                    '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($conta.Id)"
                }
                New-MgGroupMemberByRef -GroupId $grupo.Id -BodyParameter $referenciaUtilizador -ErrorAction Stop
                $membrosDoGrupo[$conta.Id] = $true
                $adicionadosAoGrupo++
                Write-Log "Adicionado: '$upn' ao grupo '$nomeGrupo'." Green
            }

            if ($Alunos -and -not [string]::IsNullOrWhiteSpace(([string]$utilizador.Turma))) {
                $mailNicknameTurma = Get-TurmaMailNickname -Ano ([string]$utilizador.Ano) -Turma ([string]$utilizador.Turma) -AcademicYear $AnoLetivo

                if (-not $equipasTurma.ContainsKey($mailNicknameTurma)) {
                    $resultadoEquipa = Get-OrCreateClassTeam -MailNickname $mailNicknameTurma
                    if ($resultadoEquipa.Created) {
                        $turmasCriadas++
                    }
                    $equipasTurma[$mailNicknameTurma] = $resultadoEquipa.Team
                }

                Add-TeamUser -GroupId $equipasTurma[$mailNicknameTurma].GroupId -User $upn -Role Member -ErrorAction Stop
                $adicionadosATurma++
                Write-Log "Adicionado: '$upn' a turma '$mailNicknameTurma'." Green
            }
        }
        catch {
            $falhados++
            Write-Log "Falha ao criar '$upn': $($_.Exception.Message)" Red
        }
    }

    Write-Log "Concluido: $processados processados, $criados criados, $existentes ja existiam, $adicionadosAoGrupo adicionados ao grupo, $jaNoGrupo ja pertenciam ao grupo, $turmasCriadas turmas criadas, $adicionadosATurma adicionados a turma, $falhados falhados." Cyan

    if ($falhados -gt 0) {
        exit 1
    }
}
catch {
    Write-Log "Erro fatal: $($_.Exception.Message)" Red
    exit 1
}