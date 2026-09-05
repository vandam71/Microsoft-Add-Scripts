# Microsoft-Add-Scripts

`adicionar_utilizadores.ps1` adiciona utilizadores a equipas do Microsoft Teams a partir de um ficheiro CSV.

## Requisitos

- PowerShell 7 ou Windows PowerShell. O script verifica o modulo `MicrosoftTeams` e pede autorizacao para o instalar para o utilizador atual quando necessario.
- Permissoes para gerir os membros das equipas de destino.
- Um CSV separado por ponto e virgula ou virgula, com as colunas `equipa` e `email`.

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

`criar_utilizadores.ps1` cria contas no Microsoft Entra ID e adiciona-as ao grupo que define as respetivas permissoes. O script verifica os modulos `Microsoft.Graph.Users` e `Microsoft.Graph.Groups` e pede autorizacao para os instalar para o utilizador atual quando necessario. Para `-Alunos`, tambem verifica o modulo `MicrosoftTeams` para associar os alunos as turmas. A conta que inicia sessao necessita das permissoes delegadas `User.ReadWrite.All` e `GroupMember.ReadWrite.All`.

O CSV aceita ponto e virgula ou virgula como delimitador e usa diretamente a exportacao da base de dados. As colunas `Processo`, `Nome` e `NIF` sao obrigatorias; `Ano` e `Turma` sao usadas para associar alunos a uma turma. Na coluna `Ano`, o script extrai apenas o primeiro numero e ignora o resto do texto (por exemplo: `1`, `1º ano`, `Ano 1`, `1 - turma D`):

```csv
Processo;Nome;NIF;Ano;Turma
7738;Tomas Pereira Coelho;288519868;5º ano;A
```

```powershell
.\criar_utilizadores.ps1 -Alunos -Ficheiro .\novos-alunos.csv
.\criar_utilizadores.ps1 -Docentes -Ficheiro .\novos-docentes.csv
```

O delimitador e detetado automaticamente pelo cabecalho. Nao misture ponto e virgula e virgula na mesma linha de cabecalho. O `upn` e criado como `Processo@alunos.amadeo.pt`. O primeiro nome torna-se o `nome`; os restantes tornam-se o `apelido`. A palavra-passe de novos alunos e `Aluno` seguido dos primeiros cinco algarismos do NIF e `#` (por exemplo, `Aluno28851#`); para docentes, o prefixo e `Docente`.

O script associa alunos ao grupo `O365-Alunos` e docentes ao grupo `O365-Professores`. Para alunos com `Turma` preenchida, tambem adiciona a conta a equipa com o `MailNickName` `<ano>-<ano><turma>-<ano-letivo>`; por exemplo, `5º ano` e `A` resultam em `5-5A-2026-2027`. Quando essa equipa ainda nao existe, o script cria-a automaticamente com o modelo Teams `EDU_Class` (Turma), equivalente a escolher `Criar uma equipa a partir de um modelo` e depois `Turma` na interface. O ano letivo e calculado automaticamente a partir da data atual e pode ser definido explicitamente:

```powershell
.\criar_utilizadores.ps1 -Alunos -Ficheiro .\novos-alunos.csv -AnoLetivo 2026-2027
```

Tambem adiciona ao grupo contas cujo UPN ja exista; nunca altera a palavra-passe de uma conta existente. As novas contas nao sao obrigadas a alterar a palavra-passe no primeiro inicio de sessao.