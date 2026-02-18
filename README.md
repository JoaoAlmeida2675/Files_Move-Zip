# Script de Organização, Movimento e Zipagem de Ficheiros

## Descrição Geral

Este script em **PowerShell** automatiza a organização, movimentação e arquivamento de ficheiros em pastas locais ou partilhas de rede.

Funcionalidades principais:

* Organiza ficheiros antigos em pastas específicas (`ARQUIVO`);
* Move ficheiros para uma estrutura centralizada;
* Cria ficheiros ZIP por ano para arquivamento eficiente.

Inclui barra de progresso, modo **Debug** e sistema de **log**.

---

## Funcionalidades

### 1. Organização Local

* Percorre pastas no `RootPath`;
* Cria pasta `ARQUIVO` se não existir;
* Move ficheiros antigos para `ARQUIVO`.

### 2. Movimento para Nova Estrutura

* Cria `NewRootPath` se não existir;
* Procura todas as pastas `ARQUIVO`;
* Move ficheiros mantendo subpastas relativas.

### 3. Criação de ZIP por Ano

* Agrupa ficheiros por `LastWriteTime`;
* Cria ZIP por ano;
* Remove ficheiros originais após confirmação do ZIP.

---

## Configuração

### Modo Debug

```powershell
$DebugMode = $true # true para ativar, false para desativar
```

### Log

```powershell
$LogPath = "CAMINHO\FICHEIRO\LOG"
$LogFile = Join-Path $LogPath "FileMover_Logs.txt"
```

### Caminhos de Trabalho

```powershell
$RootPath = "\CAMINHO\DA\PASTA\RAIZ"
$NewPath = "\CAMINHO\NOVA\RAIZ"
$NewRootPath = Join-Path $NewPath "NOVA\ESTRUTURA\PASTAS"
```

---

## Requisitos

* Windows PowerShell 5.1 ou superior;
* Acesso de escrita nas pastas de origem, destino e log;
* Assemblies: `System.Windows.Forms`, `System.IO.Compression`, `System.IO.Compression.FileSystem`.

---

## Uso

1. Abra PowerShell com permissões adequadas;
2. Ajuste caminhos e modo Debug;
3. Execute:

```powershell
.\FileOrganizer.ps1
```

4. Consulte a barra de progresso e o log.

---

## Mensagem Final

* Modo Debug: contadores detalhados de ficheiros movidos e ZIPs criados;
* Produção: confirmação de conclusão e log.

---

## Licença

Uso livre, fornecido **"as is"**.
