# ============================================================
# SCRIPT DE ORGANIZAÇÃO, MOVIMENTO E ZIPAGEM DE FICHEIROS
# ============================================================
#
# DESCRIÇÃO GERAL
# ------------------------------------------------------------
# Este script executa três processos principais para organizar
# ficheiros em pastas, movê-los para uma estrutura centralizada
# e criar ficheiros ZIP para arquivamento.
#
# ============================================================
# FASE 1 — ORGANIZAÇÃO LOCAL
# ============================================================
#
# - Percorre as pastas de primeiro nível em RootPath
# - Cria "ARQUIVO" se não existir
# - Move ficheiros cujo LastWriteTime seja de ano anterior
#
# ============================================================
# FASE 2 — MOVIMENTO PARA ESTRUTURA NOVA DE PASTAS
# ============================================================
#
# - Cria a nova estrutura se não existir
# - Procura todas as pastas ARQUIVO
# - Move ficheiros para a estrutura mantendo subpastas relativas
#
# ============================================================
# FASE 3 — CRIAÇÃO DE ZIP POR ANO
# ============================================================
#
# - Agrupa ficheiros por ano com base no LastWriteTime
# - Cria ficheiro ZIP por ano em cada subpasta
# - Remove ficheiros originais após confirmação do ZIP
#
# Inclui:
# - Barra de progresso visual
# - Modo Debug para auditoria e contagem de ficheiros
# - Sistema de Log para registar todas as operações num ficheiro
#
# Compatível com pastas locais e partilhas de rede
#
# ============================================================


# ============================================================
# MODO DEBUG
# ============================================================
#
# O modo debug permite ativar contadores e mensagens adicionais
# para verificar se todos os passos do script estão a ser
# executados corretamente. Em produção, pode ser desligado.
#
$DebugMode = $true   # Alterar para $false para desativar debug


# ============================================================
# CONFIGURAÇÃO DO LOG
# ============================================================
#
# Define o local e o ficheiro de log onde todas as ações
# serão registadas para posterior análise.
#
$LogPath = "CAMNIHO\FICHEIRO\LOG"   # Pasta onde o log será guardado
$LogFile = Join-Path $LogPath "FileMover_Logs.txt"  # Nome do ficheiro de log

# Criar a pasta de log caso não exista
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

# Função que escreve mensagens no log com timestamp
function Write-Log {
    param (
        [string]$Message
    )
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "[$TimeStamp] $Message"
}

# Registo inicial do processo
Write-Log "================ INÍCIO DO PROCESSO ================"


# ============================================================
# USAR ESTA PARTE PARA DEBUGGING
# ============================================================
#
# Este bloco inicializa variáveis de contagem de ficheiros
# e força o PowerShell a parar caso ocorra algum erro crítico
# durante a execução.
#
if ($DebugMode) {
    $ErrorActionPreference = "Stop"    # Parar execução em erro crítico
    $MovedToArchiveCount = 0           # Contador de ficheiros movidos para ARQUIVO
    $MovedToNewCount = 0           # Contador de ficheiros movidos para nova pasta
    $ZipCreatedCount = 0               # Contador de ZIPs criados
}


# ============================================================
# PERMISSÕES DE UI
# ============================================================
#
# Adiciona a capacidade de mostrar MessageBox no final
# para feedback do utilizador.
#
Add-Type -AssemblyName System.Windows.Forms


# ============================================================
# CONFIGURAÇÃO DOS CAMINHOS DE TRABALHO
# ============================================================
#
# Definir pasta raiz onde estão os ficheiros originais e
# pasta de destino para criar a estrutura nova.
#
$RootPath = "\CAMINHO\DA\PASTA\RAIZ"  # Pasta original
$NewPath = "\CAMINHO\NOVA\RAIZ"                  # Pasta base a ser criada
$NewRootPath = Join-Path $NewPath "NOVA\ESTRUTURA\PASTAS"  # Estrutura final de pastas

# Registar caminhos no log
Write-Log "RootPath definido como: $RootPath"
Write-Log "Destino das novas pastas definido como: $NewRootPath"


# ============================================================
# OBTENÇÃO DA DATA ATUAL
# ============================================================
#
# O ano atual será usado como referência para decidir
# quais ficheiros são antigos e devem ser arquivados.
#
$Now = Get-Date
$CurrentYear = $Now.Year    # Pode ser alterado consoante a necessidade

Write-Log "Ano atual do sistema: $CurrentYear"


# ============================================================
# INICIALIZAÇÃO DA BARRA DE PROGRESSO
# ============================================================
#
# Define função para atualizar visualmente o progresso
# da execução no PowerShell.
#
$TotalSteps = 1      # Passo inicial, será ajustado depois
$CurrentStep = 0     # Contador de progresso

function Update-Progress {
    param(
        [int]$Step,
        [string]$Activity,
        [string]$Status
    )

    Write-Progress `
        -Activity $Activity `
        -Status $Status `
        -PercentComplete (($Step / $TotalSteps) * 100)
}

Update-Progress -Step 0 -Activity "Iniciando script" -Status "Preparando execução..."


# ============================================================
# FASE 1 — ORGANIZAÇÃO LOCAL
# ============================================================
#
# - Percorre as pastas de primeiro nível em RootPath
# - Cria "ARQUIVO" se não existir
# - Move ficheiros cujo LastWriteTime seja de ano anterior
#
Write-Log "Início da FASE 1"

$ParentFolders = Get-ChildItem -Path $RootPath -Directory
$TotalSteps = $ParentFolders.Count * 2  # Ajuste para a barra de progresso

foreach ($Parent in $ParentFolders) {

    $SubFolders = Get-ChildItem -Path $Parent.FullName -Directory
    
    foreach ($Folder in $SubFolders) {

        # Ignorar se já for ARQUIVO
        if ($Folder.Name -ieq "ARQUIVO") {
            continue
        }
        
        # Definir caminho ARQUIVO
        $ArchivePath = Join-Path $Folder.FullName "ARQUIVO"

        # Criar pasta ARQUIVO caso não exista
        if (-not (Test-Path $ArchivePath)) {
            New-Item -ItemType Directory -Path $ArchivePath -Force | Out-Null
            Write-Log "Criada pasta ARQUIVO em: $ArchivePath"
        }
        
        # Obter ficheiros da subpasta
        $Files = Get-ChildItem -Path $Folder.FullName -File
        
        foreach ($File in $Files) {

            # Usar LastWriteTime para verificar o ano
            $FileYear = $File.LastWriteTime.Year
            
            if ($FileYear -lt $CurrentYear) {

                # Tentar mover ficheiro para ARQUIVO
                try {
                    Move-Item -Path $File.FullName -Destination $ArchivePath -Force -ErrorAction Stop
                    Write-Log "Movido para ARQUIVO: $($File.FullName)"
                    if ($DebugMode) { $MovedToArchiveCount++ }
                }
                catch {
                    Write-Log "ERRO ao mover para ARQUIVO: $($File.FullName) | $_"
                }
            }
        }
    }
    
    # Atualizar barra de progresso
    $CurrentStep++
    Update-Progress -Step $CurrentStep -Activity "Fase 1" -Status "Processando $($Parent.FullName)"
}

Write-Log "Fim da FASE 1"


# ============================================================
# FASE 2 — MOVIMENTO PARA ESTRUTURA NOVA DE PASTAS
# ============================================================
#
# - Cria a nova estrutura se não existir
# - Procura todas as pastas ARQUIVO
# - Move ficheiros para a estrutura mantendo subpastas relativas
#
Write-Log "Início da FASE 2"

New-Item -ItemType Directory -Path $NewRootPath -Force | Out-Null

do {

    # Obter todas as pastas ARQUIVO
    $ArchiveFolders = Get-ChildItem -Path $RootPath -Directory -Recurse |
                      Where-Object { $_.Name -ieq "ARQUIVO" }

    $FilesToMove = @()
    
    foreach ($Folder in $ArchiveFolders) {
        $FilesToMove += Get-ChildItem -Path $Folder.FullName -File
    }
    
    foreach ($File in $FilesToMove) {

        # Construir caminho relativo
        $RelativePath = $File.Directory.FullName.Substring($RootPath.Length).TrimStart('\')
        $RelativeClean = $RelativePath -replace '[\\/]+ARQUIVO$', ''
        $DestinationPath = Join-Path $NewRootPath $RelativeClean
        
        # Criar pasta destino caso não exista
        if (-not (Test-Path $DestinationPath)) {
            New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
        }
        
        # Tentar mover ficheiro para nova pasta
        try {
            Move-Item -Path $File.FullName -Destination $DestinationPath -Force -ErrorAction Stop
            Write-Log "Movido para nova pasta: $($File.FullName)"
            if ($DebugMode) {$MovedToNewCount++}
        }
        catch {
            Write-Log "ERRO ao mover para nova pasta: $($File.FullName) | $_"
        }
    }
    
    # Atualizar barra de progresso
    $CurrentStep++
    Update-Progress -Step $CurrentStep -Activity "Fase 2" -Status "Movendo ficheiros..."
    
} while ($FilesToMove.Count -gt 0)

Write-Log "Fim da FASE 2"


# ============================================================
# FASE 3 — CRIAÇÃO DE ZIP POR ANO
# ============================================================
#
# - Agrupa ficheiros por ano com base no LastWriteTime
# - Cria ficheiro ZIP por ano em cada subpasta
# - Remove ficheiros originais após confirmação do ZIP
#
Write-Log "Início da FASE 3"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$NewSubFolders = Get-ChildItem -Path $NewRootPath -Directory -Recurse
$TotalSteps += $NewSubFolders.Count

foreach ($Folder in $NewSubFolders) {

    # Obter ficheiros que não sejam ZIP
    $Files = Get-ChildItem -Path $Folder.FullName -File |
             Where-Object { $_.Extension -ne ".zip" }

    $FilesByYear = @{}
    
    foreach ($File in $Files) {

        $Year = $File.LastWriteTime.Year
        
        # Ignorar ficheiros do ano atual
        if ($Year -eq $CurrentYear) { continue }
        
        # Agrupar ficheiros por ano
        if (-not $FilesByYear.ContainsKey($Year)) { $FilesByYear[$Year] = @() }
        $FilesByYear[$Year] += $File
    }
    
    foreach ($Year in $FilesByYear.Keys) {

        $ZipPath = Join-Path $Folder.FullName ("$Year.zip")
        
        # Ignorar se ZIP já existir
        if (Test-Path $ZipPath) {
            Write-Log "ZIP já existente: $ZipPath"
            continue
        }
        
        try {
            # Criar ZIP usando Compress-Archive nativo
            $FilesToZip = $FilesByYear[$Year] | ForEach-Object { $_.FullName }

            Compress-Archive `
                -Path $FilesToZip `
                -DestinationPath $ZipPath `
                -CompressionLevel Optimal `
                -ErrorAction Stop

            # Remover ficheiros originais apenas após sucesso
            if (Test-Path $ZipPath) {
                Write-Log "ZIP criado com sucesso: $ZipPath"
                if ($DebugMode) { $ZipCreatedCount++ }

                foreach ($F in $FilesByYear[$Year]) {
                    Remove-Item -Path $F.FullName -Force -ErrorAction Stop
                }
            }

        }
        catch {
            Write-Log "ERRO ao criar ZIP $ZipPath | $_"
        }
    }
    
    # Atualizar barra de progresso
    $CurrentStep++
    Update-Progress -Step $CurrentStep -Activity "Fase 3" -Status "Processando $($Folder.FullName)"
}

Write-Log "Fim da FASE 3"


# ============================================================
# FINALIZAÇÃO
# ============================================================
#
# - Regista o fim do processo no log
# - Mostra resumo ou mensagem final dependendo do debug
#
Write-Log "================ FIM DO PROCESSO ================"

Start-Sleep -Milliseconds 500

if ($DebugMode) {

    # Mensagem de resumo detalhado quando debug ativo
    $Resumo = @"
Processo terminado.

Movidos para ARQUIVO: $MovedToArchiveCount
Movidos para Nova Pasta: $MovedToNewCount
ZIPs criados: $ZipCreatedCount

Consultar log em:
$LogFile
"@

    [System.Windows.Forms.MessageBox]::Show(
        $Resumo,
        "Resumo da Execução (DEBUG)",
        0,
        64
    )
}
else {
    # Mensagem padrão em produção
    [System.Windows.Forms.MessageBox]::Show(
        "Processo concluído.`nConsultar log em:`n$LogFile",
        "Concluído",
        0,
        64
    )
}
