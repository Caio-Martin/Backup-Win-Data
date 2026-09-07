#Requires -RunAsAdministrator
<#
Backup manual (rodar antes de formatar a maquina):
  - Fecha o Outlook se estiver aberto (libera os PST para copia)
  - PST de todos os perfis locais
  - Assinaturas/perfil do Outlook
  - Bookmarks de Chrome/Edge/Firefox de todos os perfis
  - Pastas customizadas (edite $CustomPaths abaixo)
  - Copia tudo para \\<SambaServer>\<ShareName>\<HOSTNAME>\
  - Valida cada arquivo copiado por hash SHA256 (fonte x destino)
  - Tira um print da tela como evidencia de execucao

Idempotente: pode rodar quantas vezes quiser. Arquivos ja copiados e
identicos (mesmo tamanho/data) sao pulados; so o que mudou e recopiado
e revalidado por hash.

Uso:
  .\Backup-UserData.ps1 -SambaServer 10.0.0.5 -ShareName backups
  (vai pedir usuario/senha do Samba interativamente)
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SambaServer,

    [Parameter(Mandatory = $true)]
    [string]$ShareName,

    [switch]$DeepPstScan  # varre o perfil inteiro atras de .pst, alem dos locais padrao (mais lento)
)

$ErrorActionPreference = "Stop"

# Pastas extras que voce quer incluir no backup, alem do que ja e coletado automaticamente
$CustomPaths = @(
    # "C:\Users\fulano\Documents\Projetos"
)

$Hostname       = $env:COMPUTERNAME
$Timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$SharePath      = "\\$SambaServer\$ShareName"
$DestRoot       = Join-Path $SharePath $Hostname
$DetailLogFile  = Join-Path $DestRoot "log_completo_$Timestamp.txt"
$SummaryLogFile = Join-Path $DestRoot "log_resumo_$Timestamp.txt"

$script:StartTime   = Get-Date
$script:TotalCopied = 0
$script:TotalErrors = 0

function Write-Log {
    param([string]$Message, [switch]$Summary)
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message
    Write-Host $line
    Add-Content -Path $DetailLogFile -Value $line -ErrorAction SilentlyContinue
    if ($Summary) { Add-Content -Path $SummaryLogFile -Value $line -ErrorAction SilentlyContinue }
}

function Get-RobocopyStatus {
    param([int]$Code)
    if ($Code -ge 8) { return "ERRO" }
    elseif ($Code -band 1) { return "COPIADO" }
    else { return "SEM_MUDANCAS" }
}

function Test-FileIntegrity {
    param([string]$SourceFile, [string]$DestFile)
    try {
        $srcHash = (Get-FileHash -Path $SourceFile -Algorithm SHA256 -ErrorAction Stop).Hash
        $dstHash = (Get-FileHash -Path $DestFile -Algorithm SHA256 -ErrorAction Stop).Hash
        return $srcHash -eq $dstHash
    }
    catch { return $false }
}

function Test-CopyIntegrity {
    param([string]$Source, [string]$Destination)
    if (Test-Path $Source -PathType Leaf) {
        return Test-FileIntegrity $Source $Destination
    }
    $srcFiles = Get-ChildItem $Source -Recurse -File -ErrorAction SilentlyContinue
    foreach ($f in $srcFiles) {
        $rel = $f.FullName.Substring($Source.Length).TrimStart('\')
        $destFile = Join-Path $Destination $rel
        if (-not (Test-FileIntegrity $f.FullName $destFile)) { return $false }
    }
    return $true
}

function Stop-OutlookIfRunning {
    $procs = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    if (-not $procs) {
        Write-Log "Outlook nao esta em execucao."
        return
    }
    Write-Log "Outlook esta aberto - fechando antes do backup para liberar os arquivos PST..." -Summary
    foreach ($p in $procs) { $p.CloseMainWindow() | Out-Null }

    $waited = 0
    while ((Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) -and $waited -lt 15) {
        Start-Sleep -Seconds 1
        $waited++
    }

    $remaining = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    if ($remaining) {
        Write-Log "Outlook nao fechou normalmente em 15s, forcando encerramento..." -Summary
        $remaining | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Write-Log "Outlook encerrado." -Summary
}

function Copy-IfExists {
    param([string]$Source, [string]$Destination, [string]$Label)
    if (-not (Test-Path $Source)) {
        Write-Log "SKIP - $Label nao encontrado ($Source)"
        return
    }
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    robocopy $Source $Destination /E /R:2 /W:2 /NP /LOG+:$DetailLogFile | Out-Null
    $status = Get-RobocopyStatus $LASTEXITCODE
    switch ($status) {
        "COPIADO" {
            if (Test-CopyIntegrity $Source $Destination) {
                $script:TotalCopied++
                Write-Log "OK - $Label atualizado e validado ($Source)" -Summary
            }
            else {
                $script:TotalErrors++
                Write-Log "ERRO - $Label copiado mas FALHOU na validacao pos-copia ($Source)" -Summary
            }
        }
        "SEM_MUDANCAS" { Write-Log "OK - $Label ja estava atualizado ($Source)" }
        "ERRO" {
            $script:TotalErrors++
            Write-Log "ERRO - $Label falhou, codigo robocopy $LASTEXITCODE ($Source)" -Summary
        }
    }
}

function Copy-PstFile {
    param([System.IO.FileInfo]$SourceFile, [string]$DestDir)
    New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    $destFile = Join-Path $DestDir $SourceFile.Name
    if (Test-Path $destFile) {
        $d = Get-Item $destFile
        if ($d.Length -eq $SourceFile.Length -and $d.LastWriteTime -eq $SourceFile.LastWriteTime) {
            return "SEM_MUDANCAS"
        }
    }
    try {
        Copy-Item $SourceFile.FullName -Destination $destFile -Force -ErrorAction Stop
    }
    catch {
        Write-Log "ERRO - falha ao copiar PST $($SourceFile.FullName): $_"
        return "ERRO"
    }
    if (Test-FileIntegrity $SourceFile.FullName $destFile) {
        return "COPIADO"
    }
    Write-Log "ERRO - validacao pos-copia falhou para PST $($SourceFile.FullName) (hash nao confere)"
    return "ERRO_VALIDACAO"
}

function Backup-User {
    param([System.IO.DirectoryInfo]$UserProfile)

    $user = $UserProfile.Name
    $userDest = Join-Path $DestRoot $user
    Write-Log "Processando usuario: $user"

    # PST - locais padrao
    $pstDefaultPaths = @(
        "$($UserProfile.FullName)\AppData\Local\Microsoft\Outlook",
        "$($UserProfile.FullName)\Documents\Outlook Files"
    )
    $pstFiles = @()
    foreach ($p in $pstDefaultPaths) {
        if (Test-Path $p) { $pstFiles += Get-ChildItem $p -Filter *.pst -File -ErrorAction SilentlyContinue }
    }
    if ($DeepPstScan) {
        $pstFiles += Get-ChildItem $UserProfile.FullName -Recurse -Filter *.pst -File -ErrorAction SilentlyContinue
    }
    $pstFiles = $pstFiles | Sort-Object FullName -Unique

    if ($pstFiles.Count -gt 0) {
        $pstDest = Join-Path $userDest "Outlook_PST"
        $copied = 0; $unchanged = 0; $errors = 0
        foreach ($pst in $pstFiles) {
            switch (Copy-PstFile $pst $pstDest) {
                "COPIADO"        { $copied++ }
                "SEM_MUDANCAS"   { $unchanged++ }
                "ERRO"           { $errors++ }
                "ERRO_VALIDACAO" { $errors++ }
            }
        }
        $script:TotalCopied += $copied
        $script:TotalErrors += $errors
        Write-Log "PST ($user): $copied copiado(s), $unchanged ja atualizado(s), $errors erro(s)" -Summary
    }
    else {
        Write-Log "SKIP - nenhum PST encontrado para $user" -Summary
    }

    Copy-IfExists "$($UserProfile.FullName)\AppData\Roaming\Microsoft\Signatures" (Join-Path $userDest "Outlook_Signatures") "Assinaturas Outlook ($user)"
    Copy-IfExists "$($UserProfile.FullName)\AppData\Local\Google\Chrome\User Data\Default\Bookmarks" (Join-Path $userDest "Chrome") "Bookmarks Chrome ($user)"
    Copy-IfExists "$($UserProfile.FullName)\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks" (Join-Path $userDest "Edge") "Bookmarks Edge ($user)"

    Copy-IfExists "$($UserProfile.FullName)\Desktop"   (Join-Path $userDest "Desktop")   "Desktop ($user)"
    Copy-IfExists "$($UserProfile.FullName)\Downloads" (Join-Path $userDest "Downloads") "Downloads ($user)"
    Copy-IfExists "$($UserProfile.FullName)\Documents" (Join-Path $userDest "Documents") "Documents ($user)"

    $ffProfilesPath = "$($UserProfile.FullName)\AppData\Roaming\Mozilla\Firefox\Profiles"
    if (Test-Path $ffProfilesPath) {
        $ffProfiles = Get-ChildItem $ffProfilesPath -Directory -Filter "*.default*" -ErrorAction SilentlyContinue
        foreach ($ffp in $ffProfiles) {
            Copy-IfExists "$($ffp.FullName)\places.sqlite" (Join-Path $userDest "Firefox\$($ffp.Name)") "Bookmarks Firefox ($user/$($ffp.Name))"
        }
    }
}

# --- Autenticacao no servidor Samba ---
$cred = Get-Credential -Message "Credenciais para acessar $SharePath"

# Remove qualquer mapeamento antigo/travado antes de conectar (torna o script re-executavel)
net use $SharePath /delete /y 2>&1 | Out-Null

net use $SharePath $cred.GetNetworkCredential().Password /user:$($cred.UserName) | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Falha ao autenticar em $SharePath (net use retornou codigo $LASTEXITCODE)"
}

try {
    New-Item -ItemType Directory -Path $DestRoot -Force | Out-Null
    Write-Log "Backup iniciado - host: $Hostname -> destino: $DestRoot" -Summary

    Stop-OutlookIfRunning

    $excluded = @("Public", "Default", "Default User", "All Users")
    $userProfiles = Get-ChildItem "C:\Users" -Directory | Where-Object { $excluded -notcontains $_.Name }

    foreach ($userProfile in $userProfiles) {
        Backup-User $userProfile
    }

    foreach ($path in $CustomPaths) {
        $name = Split-Path $path -Leaf
        Copy-IfExists $path (Join-Path $DestRoot "Custom\$name") "Pasta customizada ($path)"
    }

    Write-Log "Copia de arquivos concluida."

    # --- Print de evidencia da execucao ---
    # Em try/catch proprio: maquinas sem sessao grafica (ex: Windows Server Core)
    # podem nao ter System.Windows.Forms/System.Drawing disponiveis, e isso nao
    # pode derrubar um backup que ja copiou os arquivos com sucesso.
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop

        $bounds = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)

        $screenshotPath = Join-Path $DestRoot "evidencia_backup_${Hostname}_${Timestamp}.png"
        $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose()
        $bitmap.Dispose()

        Write-Log "Print de evidencia salvo em: $screenshotPath" -Summary
    }
    catch {
        $script:TotalErrors++
        Write-Log "AVISO - nao foi possivel gerar o print de evidencia (sem sessao grafica?): $_" -Summary
    }

    $duration = (Get-Date) - $script:StartTime
    Write-Log ("RESUMO: {0} item(ns) copiado/atualizado, {1} erro(s), duracao {2:mm\:ss}" -f $script:TotalCopied, $script:TotalErrors, $duration) -Summary
    Write-Log "Backup finalizado com sucesso." -Summary

    Write-Host "`nBackup completo. Destino: $DestRoot" -ForegroundColor Green
    Write-Host "Log completo: $DetailLogFile"
    Write-Host "Log resumo:   $SummaryLogFile"
}
catch {
    Write-Log "ERRO FATAL: $_" -Summary
    throw
}
finally {
    net use $SharePath /delete /y 2>&1 | Out-Null
}
