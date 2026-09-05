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

if (-not (Test-Path -LiteralPath $Ficheiro -PathType Leaf)) {
    throw "O ficheiro CSV '$Ficheiro' nao foi encontrado."
}

$tipoUtilizador = if ($Alunos) { 'alunos' } else { 'docentes' }
$role = if ($Docentes) { 'Owner' } else { 'Member' }

try {
    Ensure-Module -Name MicrosoftTeams

    Write-Log 'A ligar ao Microsoft Teams...' Cyan
    Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
    Write-Log "Ligacao estabelecida. A adicionar $tipoUtilizador como $role." Green

    $utilizadores = @(Import-UsersCsv -Path $Ficheiro)
    if ($utilizadores.Count -eq 0) {
        throw "O ficheiro CSV '$Ficheiro' nao contem registos."
    }

    $processados = 0
    $adicionados = 0
    $falhados = 0

    foreach ($utilizador in $utilizadores) {
        $processados++
        $equipa = ([string]$utilizador.equipa).Trim()
        $email = ([string]$utilizador.email).Trim()

        if ([string]::IsNullOrWhiteSpace($equipa) -or [string]::IsNullOrWhiteSpace($email)) {
            $falhados++
            Write-Log "Registo $processados ignorado: as colunas 'equipa' e 'email' sao obrigatorias." Yellow
            continue
        }

        try {
            Write-Log "[$processados/$($utilizadores.Count)] A procurar a equipa '$equipa' para '$email'..."
            $equipas = @(Get-Team -MailNickName $equipa -ErrorAction Stop)

            if ($equipas.Count -eq 0) {
                throw "Nao foi encontrada nenhuma equipa com MailNickName '$equipa'."
            }

            if ($equipas.Count -gt 1) {
                throw "Foram encontradas $($equipas.Count) equipas com MailNickName '$equipa'."
            }

            Add-TeamUser -GroupId $equipas[0].GroupId -User $email -Role $role -ErrorAction Stop
            $adicionados++
            Write-Log "Adicionado: '$email' como $role na equipa '$equipa'." Green
        }
        catch {
            $falhados++
            Write-Log "Falha ao adicionar '$email' na equipa '$equipa': $($_.Exception.Message)" Red
        }
    }

    Write-Log "Concluido: $processados processados, $adicionados adicionados, $falhados falhados." Cyan

    if ($falhados -gt 0) {
        exit 1
    }
}
catch {
    Write-Log "Erro fatal: $($_.Exception.Message)" Red
    exit 1
}