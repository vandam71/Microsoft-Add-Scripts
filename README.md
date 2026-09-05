# Microsoft-Add-Scripts

`adicionar_utilizadores.ps1` adiciona utilizadores a equipas do Microsoft Teams a partir de um ficheiro CSV.

## Requisitos

- PowerShell 7 ou Windows PowerShell com o modulo `MicrosoftTeams` instalado.
- Permissoes para gerir os membros das equipas de destino.
- Um CSV separado por ponto e virgula, com as colunas `equipa` e `email`.

Exemplo de CSV:

```csv
equipa;email
matematica-10a;aluno@example.org
matematica-10a;docente@example.org
```

## Utilizacao

Adiciona alunos como membros:

```powershell
.\adicionar_utilizadores.ps1 -Alunos -Ficheiro .\alunos.csv
```

Adiciona docentes como proprietarios:

```powershell
.\adicionar_utilizadores.ps1 -Docentes -Ficheiro .\docentes.csv
```

Escolha apenas uma das opcoes: `-Alunos` ou `-Docentes`. O script mostra o progresso por registo, continua quando um registo falha e termina com codigo `1` se existir alguma falha.