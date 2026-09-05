[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Equipa
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

try {
    Ensure-Module -Name MicrosoftTeams

    Write-Log 'A ligar ao Microsoft Teams...' Cyan
    Connect-MicrosoftTeams -ErrorAction Stop | Out-Null

    Write-Log "A procurar a equipa '$Equipa'..." Cyan
    $equipas = @(Get-Team -MailNickName $Equipa -ErrorAction Stop)
    if ($equipas.Count -eq 0) {
        throw "Nao foi encontrada nenhuma equipa com MailNickName '$Equipa'."
    }

    if ($equipas.Count -gt 1) {
        throw "Foram encontradas $($equipas.Count) equipas com MailNickName '$Equipa'."
    }

    $equipaEncontrada = $equipas[0]
    $membros = @(Get-TeamUser -GroupId $equipaEncontrada.GroupId -ErrorAction Stop)
    if ($membros.Count -eq 0) {
        Write-Log "A equipa '$Equipa' ja nao tem membros." Yellow
        exit 0
    }

    $processados = 0
    $removidos = 0
    $falhados = 0

    foreach ($membro in $membros) {
        $processados++
        $email = ([string]$membro.User).Trim()
        $role = ([string]$membro.Role).Trim()

        if ([string]::IsNullOrWhiteSpace($email) -or [string]::IsNullOrWhiteSpace($role)) {
            $falhados++
            Write-Log "Registo $processados ignorado: dados de utilizador incompletos." Yellow
            continue
        }

        try {
            Remove-TeamUser -GroupId $equipaEncontrada.GroupId -User $email -Role $role -ErrorAction Stop
            $removidos++
            Write-Log "Removido: '$email' ($role)." Green
        }
        catch {
            $falhados++
            Write-Log "Falha ao remover '$email' ($role): $($_.Exception.Message)" Red
        }
    }

    Write-Log "Concluido: $processados processados, $removidos removidos, $falhados falhados." Cyan
    if ($falhados -gt 0) {
        exit 1
    }
}
catch {
    Write-Log "Erro fatal: $($_.Exception.Message)" Red
    exit 1
}
