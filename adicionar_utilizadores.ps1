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
$role = if ($Docentes) { 'Owner' } else { 'Member' }

try {
    Write-Log 'A importar o modulo MicrosoftTeams...' Cyan
    Import-Module MicrosoftTeams -ErrorAction Stop

    Write-Log 'A ligar ao Microsoft Teams...' Cyan
    Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
    Write-Log "Ligacao estabelecida. A adicionar $tipoUtilizador como $role." Green

    $utilizadores = @(Import-Csv -LiteralPath $Ficheiro -Delimiter ';' -ErrorAction Stop)
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