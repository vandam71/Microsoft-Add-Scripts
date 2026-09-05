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

## Criar utilizadores

`criar_utilizadores.ps1` cria contas no Microsoft Entra ID e adiciona-as ao grupo que define as respetivas permissoes. Requer os modulos `Microsoft.Graph.Users` e `Microsoft.Graph.Groups`, e as permissoes delegadas `User.ReadWrite.All` e `GroupMember.ReadWrite.All` para a conta que inicia sessao.

O CSV tambem e separado por ponto e virgula e requer as colunas `nome`, `apelido`, `upn` e `password`:

```csv
nome;apelido;upn;password
Ana;Silva;ana.silva@contoso.onmicrosoft.com;PasswordTemporaria123!
```

```powershell
.\criar_utilizadores.ps1 -Alunos -Ficheiro .\novos-alunos.csv
.\criar_utilizadores.ps1 -Docentes -Ficheiro .\novos-docentes.csv
```

O `upn` e o endereco utilizado para iniciar sessao, normalmente o email do utilizador, e tem de usar um dominio verificado no tenant. O script associa alunos ao grupo `O365-Alunos` e docentes ao grupo `O365-Professores`. Tambem adiciona ao grupo contas cujo UPN ja exista; nunca altera a palavra-passe de uma conta existente. As novas contas mantem a palavra-passe indicada no CSV e nao sao obrigadas a altera-la no primeiro inicio de sessao. O CSV contem palavras-passe, por isso deve ser guardado, partilhado e removido de acordo com as regras de seguranca da organizacao.