# Backup-UserData.ps1

Script PowerShell para backup manual de dados de usuário em máquinas Windows, pensado para ser executado **antes de formatar** um computador. Copia os dados para um compartilhamento de rede (Samba/SMB) e valida cada arquivo copiado por hash.

## O que o script faz

1. **Fecha o Outlook**, se estiver aberto, para liberar os arquivos `.pst` para cópia (tenta fechar normalmente por até 15s, depois força o encerramento do processo).
2. Percorre **todos os perfis de usuário locais** em `C:\Users` (excluindo `Public`, `Default`, `Default User` e `All Users`) e, para cada um, copia:
   - **Arquivos PST do Outlook**
     - Locais padrão verificados: `AppData\Local\Microsoft\Outlook` e `Documents\Outlook Files`.
     - Com `-DeepPstScan`, também varre o perfil inteiro em busca de `.pst` (mais lento).
   - **Assinaturas do Outlook** (`AppData\Roaming\Microsoft\Signatures`).
   - **Bookmarks do Chrome** (`AppData\Local\Google\Chrome\User Data\Default\Bookmarks`).
   - **Bookmarks do Edge** (`AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks`).
   - **Bookmarks do Firefox** (`places.sqlite` de cada perfil `*.default*` em `AppData\Roaming\Mozilla\Firefox\Profiles`).
   - **Desktop** (`Desktop`).
   - **Downloads** (`Downloads`).
   - **Documents** (`Documents`).
3. Copia **pastas customizadas** adicionais definidas manualmente na variável `$CustomPaths` no início do script (por padrão, vazia).
4. Copia tudo para `\\<SambaServer>\<ShareName>\<HOSTNAME>\`, organizado por usuário.
5. **Valida cada arquivo copiado por hash SHA256** (compara origem x destino) — se a validação falhar, é registrado como erro.
6. **Tira um print da tela inteira** (todos os monitores) como evidência de execução, salvo como PNG no destino.
7. Gera **dois logs**: um completo (`log_completo_<timestamp>.txt`, todas as mensagens) e um resumo (`log_resumo_<timestamp>.txt`, apenas eventos importantes), ambos no destino da rede.

## Idempotência

O script pode ser executado quantas vezes forem necessárias:

- **Arquivos/pastas em geral** são copiados via `robocopy /E` — o robocopy só recopia o que mudou (novo, tamanho ou data diferentes).
- **Arquivos PST** são comparados por tamanho + data de modificação antes de decidir se recopia.
- Em ambos os casos, todo arquivo efetivamente copiado é **revalidado por hash SHA256** contra a origem.

## Requisitos

- Deve ser executado **como Administrador** (`#Requires -RunAsAdministrator`).
- PowerShell no Windows.
- Acesso de rede a um servidor Samba/SMB com um compartilhamento (`ShareName`) acessível por usuário/senha.
- Credenciais válidas para o compartilhamento (solicitadas interativamente via `Get-Credential` durante a execução).

## Uso

```powershell
.\Backup-UserData.ps1 -SambaServer 10.0.0.5 -ShareName backups
```

O script pedirá usuário/senha do Samba interativamente (via `Get-Credential`).

### Parâmetros

| Parâmetro       | Obrigatório | Descrição                                                                                     |
|-----------------|:-----------:|------------------------------------------------------------------------------------------------|
| `-SambaServer`  | Sim         | Endereço/hostname do servidor Samba de destino.                                               |
| `-ShareName`    | Sim         | Nome do compartilhamento no servidor Samba.                                                   |
| `-DeepPstScan`  | Não         | Varre o perfil inteiro do usuário atrás de `.pst`, além dos locais padrão (mais lento).        |

### Personalizando pastas extras

Para incluir pastas adicionais no backup, edite a variável `$CustomPaths` no topo do script:

```powershell
$CustomPaths = @(
    "C:\Users\fulano\Documents\Projetos"
)
```

Cada caminho listado é copiado para `\\<SambaServer>\<ShareName>\<HOSTNAME>\Custom\<NomeDaPasta>`.

## Estrutura do destino gerado

```
\\<SambaServer>\<ShareName>\<HOSTNAME>\
├── log_completo_<timestamp>.txt
├── log_resumo_<timestamp>.txt
├── evidencia_backup_<HOSTNAME>_<timestamp>.png
├── Custom\
│   └── <NomeDaPasta>\           (pastas de $CustomPaths)
└── <usuario>\
    ├── Outlook_PST\             (arquivos .pst)
    ├── Outlook_Signatures\      (assinaturas do Outlook)
    ├── Chrome\Bookmarks
    ├── Edge\Bookmarks
    ├── Firefox\<perfil>\places.sqlite
    ├── Desktop\
    ├── Downloads\
    └── Documents\
```

## Logs

- **Log completo**: todas as mensagens (inclusive `SKIP` de itens não encontrados, saída detalhada).
- **Log resumo**: apenas eventos relevantes (início/fim do backup, Outlook fechado, itens copiados/atualizados, erros, resumo final com total de itens copiados, erros e duração).

Ao final, o script também imprime no console o caminho do destino e dos dois arquivos de log.

## Tratamento de erros e casos especiais

- Se um item esperado (pasta/arquivo) não existir na origem, é registrado como `SKIP` e o backup continua normalmente.
- Se a cópia de um PST falhar ou a validação de hash pós-cópia falhar, é contabilizado como erro, mas o backup dos demais itens/usuários continua.
- Se a máquina não tiver sessão gráfica disponível (ex.: Windows Server Core), a etapa de print de evidência falha de forma isolada (contabilizada como erro/aviso), sem interromper o restante do backup, já que os arquivos já foram copiados com sucesso.
- Qualquer mapeamento de rede antigo/travado para o compartilhamento é removido (`net use ... /delete`) antes de conectar, tornando o script re-executável sem exigir limpeza manual.
- Em caso de erro fatal (ex.: falha de autenticação no Samba), o script lança exceção e interrompe a execução; o mapeamento de rede é sempre desfeito ao final (bloco `finally`), com sucesso ou falha.

## Segurança

- As credenciais do Samba são solicitadas interativamente (`Get-Credential`) e usadas apenas em memória para `net use`; não são persistidas em disco pelo script.
- Recomenda-se usar uma conta com permissão apenas de escrita no compartilhamento de destino.
