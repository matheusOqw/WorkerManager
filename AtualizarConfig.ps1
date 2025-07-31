Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Define caminho do log (salvo na mesma pasta do script)
$Global:LogPath = Join-Path $PSScriptRoot "log_modificacoes_worker.txt"

Function Registrar-Log {
    param (
        [string]$Servico,
        [string]$Parametro,
        [string]$ValorAntigo,
        [string]$ValorNovo,
        [string]$ArquivoConfig
    )

    $log = "[{0}] - {1}: {2} -> {3} | Serviço: {4} | Arquivo: {5}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Parametro, $ValorAntigo, $ValorNovo, $Servico, $ArquivoConfig
    Add-Content -Path $Global:LogPath -Value $log
}

# Função para listar os serviços Benner Tasks Worker validando qual fila está sendo lida pelo arquivo config (Z_BTLREQUISICOES ou Z_AGENDAMENTOREQUISICOES) =====================================
Function Selecionar-ServicosInterativamente {
    param (
        [Parameter(Mandatory = $true)]
        [array]$ListaDeServicos
    )

    if (-not $ListaDeServicos -or $ListaDeServicos.Count -eq 0) {
        Write-Warning "Lista de serviços vazia."
        return @()
    }

    Write-Host "`n📋 Serviços encontrados:`n"
    for ($i = 0; $i -lt $ListaDeServicos.Count; $i++) {
        Write-Host "$($i + 1)) $($ListaDeServicos[$i].DisplayName)"
    }

    $opcao = Read-Host "`nDigite os números dos serviços separados por vírgula (ou 'todos' para alterar todos)"
    $selecionados = @()

    if ($opcao.Trim().ToLower() -eq "todos") {
        $selecionados = $ListaDeServicos
    } else {
        $indices = @()
        $opcao.Split(",") | ForEach-Object {
            $numeroLimpo = $_.Trim()
            $parsed = 0
            if ([int]::TryParse($numeroLimpo, [ref]$parsed)) {
                if ($parsed -gt 0 -and $parsed -le $ListaDeServicos.Count) {
                    $indices += ($parsed - 1)
                }
            }
        }
        $selecionados = $indices | ForEach-Object { $ListaDeServicos[$_] }
    }

    if (-not $selecionados -or $selecionados.Count -eq 0) {
        Write-Warning "Nenhum serviço válido selecionado."
    }

    return $selecionados
}
Function Obter-ServicosPorFila {
    param (
        [string]$tipoFiltro
    )

    $mapaFila = @{
        "1" = "Z_AGENDAMENTOREQUISICOES"
        "2" = "Z_BTLREQUISICOES"
    }

    $servicos = Get-WmiObject Win32_Service | Where-Object {
        $_.DisplayName -like "Benner Tasks Worker*"
    }

    if (-not $servicos.Count) {
        Write-Host "❌ Nenhum serviço encontrado com nome 'Benner Tasks Worker*'."
        return @()
    }

    $filtrados = @()

    foreach ($servico in $servicos) {
        $caminhoExe = $servico.PathName.Trim('"')
        $pasta = Split-Path $caminhoExe -Parent
        $configPath = Join-Path $pasta "Benner.Tecnologia.Tasks.Worker.exe.config"

        if (Test-Path $configPath) {
            try {
                [xml]$xml = Get-Content $configPath
                $fila = $xml.configuration.appSettings.add |
                    Where-Object { $_.key -eq "Fila" } |
                    Select-Object -ExpandProperty value

                if ($tipoFiltro -eq "3" -or $fila -eq $mapaFila[$tipoFiltro]) {
                    $filtrados += $servico
                }
            } catch {
                Write-Warning "⚠️ Erro ao ler config de '$($servico.DisplayName)'"
            }
        } else {
            Write-Warning "⚠️ Arquivo .config não encontrado para '$($servico.DisplayName)'"
        }
    }

    if ($filtrados.Count -eq 0) {
        switch ($tipoFiltro) {
            "1" { Write-Host "`n❌ Nenhum Worker encontrado lendo a fila Z_AGENDAMENTOREQUISICOES." }
            "2" { Write-Host "`n❌ Nenhum Worker encontrado lendo a fila Z_BTLREQUISICOES." }
            "3" { Write-Host "`n❌ Nenhum serviço encontrado com qualquer fila configurada." }
            default { Write-Host "`n❌ Valor inválido." }
        }
    }

    return $filtrados
}

# Função para atualizar número de provider no worker ================================================================================================
Function Atualizar-NumeroDeProviders {
    do {
        $NovoValor = Read-Host "Informe o novo valor para 'NumeroDeProviders' (somente números)"
        $valido = $NovoValor -match '^[0-9]+$'
        if (-not $valido) { Write-Warning "Digite apenas números." }
    } while (-not $valido)

    Write-Host "`nDeseja alterar os serviços de:"
    Write-Host "1 - Apenas Agendamento"
    Write-Host "2 - Apenas BTL"
    Write-Host "3 - Todos"
    $tipoFiltro = Read-Host "Digite 1, 2 ou 3"

    $servicosFiltrados = Obter-ServicosPorFila -tipoFiltro $tipoFiltro

    if (-not $servicosFiltrados -or $servicosFiltrados.Count -eq 0) {
        return
    }

    Write-Host "`n📋 Serviços encontrados:`n"
    for ($i = 0; $i -lt $servicosFiltrados.Count; $i++) {
        Write-Host "$($i + 1)) $($servicosFiltrados[$i].DisplayName)"
    }

    $selecionados = @()
    $opcao = Read-Host "`nDigite os números dos serviços separados por vírgula (ou 'todos' para alterar todos)"

    if ($opcao.Trim().ToLower() -eq "todos") {
        $selecionados = $servicosFiltrados
    } else {
        $indices = @()
        $opcao.Split(",") | ForEach-Object {
            $numeroLimpo = $_.Trim()
            $parsed = 0
            if ([int]::TryParse($numeroLimpo, [ref]$parsed)) {
                if ($parsed -gt 0 -and $parsed -le $servicosFiltrados.Count) {
                    $indices += ($parsed - 1)
                }
            }
        }
        $selecionados = $indices | ForEach-Object { $servicosFiltrados[$_] }
    }

    if (-not $selecionados -or $selecionados.Count -eq 0) {
        Write-Warning "Nenhum serviço válido selecionado."
        return
    }

    foreach ($servico in $selecionados) {
        $caminhoExe = $servico.PathName -replace '"', ''
        $diretorio = Split-Path $caminhoExe
        $configPath = Join-Path $diretorio "Benner.Tecnologia.Tasks.Worker.exe.config"

        if (-not (Test-Path $configPath)) {
            Write-Warning "⚠️ Arquivo não encontrado: $configPath"
            continue
        }

        try {
            [xml]$xml = Get-Content $configPath
            $node = $xml.configuration.appSettings.add | Where-Object { $_.key -eq "NumeroDeProviders" }

            if ($node) {
                if ($node.value -ne $NovoValor) {
                    $valorAntigo = $node.value
                    $node.value = "$NovoValor"
                    $xml.Save($configPath)
                    Write-Host "✅ $($servico.DisplayName): $valorAntigo → $NovoValor"
                    Registrar-Log -Servico $servico.DisplayName -Parametro "NumeroDeProviders" -ValorAntigo $valorAntigo -ValorNovo $NovoValor -ArquivoConfig $configPath

                    Write-Host "🔄 Reiniciando o serviço '$($servico.DisplayName)'..."
                    try {
                        Stop-Service -Name $servico.Name -Force -ErrorAction Stop
                        Start-Service -Name $servico.Name -ErrorAction Stop
                        Write-Host "✅ Serviço '$($servico.DisplayName)' reiniciado com sucesso." -ForegroundColor Green
                    } catch {
                        Write-Warning "❌ Erro ao reiniciar o serviço '$($servico.Name)': $_"
                    }
                } else {
                    Write-Host "ℹ️ $($servico.DisplayName) já possui o valor desejado ($NovoValor)" -ForegroundColor Yellow
                }
            } else {
                Write-Warning "⚠️ Parâmetro 'NumeroDeProviders' não encontrado."
            }

        } catch {
            Write-Warning "❌ Erro ao processar ${configPath}: $_"
        }
    }

    Write-Host "`n✔️ Atualização concluída." -ForegroundColor Green
}

# Função para atualizar parâmetro COMFree no worker =================================================================================================
Function Atualizar-UseCOMFree {
    do {
        $NovoValor = Read-Host "Informe o novo valor para 'UseCOMFree' (true ou false)"
        $NovoValor = $NovoValor.Trim().ToLower()
        $valido = $NovoValor -in @("true", "false")
        if (-not $valido) { Write-Warning "Digite apenas 'true' ou 'false'." }
    } while (-not $valido)

    # Selecionar tipo de serviço com base na fila
    Write-Host "`nDeseja alterar os serviços de:"
    Write-Host "1 - Apenas Agendamento"
    Write-Host "2 - Apenas BTL"
    Write-Host "3 - Todos"
    $tipoFiltro = Read-Host "Digite 1, 2 ou 3"

    # Obter serviços com base na Fila do .config
    $servicosFiltrados = Obter-ServicosPorFila -tipoFiltro $tipoFiltro

    if (-not $servicosFiltrados -or $servicosFiltrados.Count -eq 0) {
        return
    }

    Write-Host "`n📋 Serviços encontrados:`n"
    for ($i = 0; $i -lt $servicosFiltrados.Count; $i++) {
        Write-Host "$($i + 1)) $($servicosFiltrados[$i].DisplayName)"
    }

    $selecionados = @()
    $opcao = Read-Host "`nDigite os números dos serviços separados por vírgula (ou 'todos' para alterar todos)"

    if ($opcao.Trim().ToLower() -eq "todos") {
        $selecionados = $servicosFiltrados
    } else {
        $indices = @()
        $opcao.Split(",") | ForEach-Object {
            $numeroLimpo = $_.Trim()
            $parsed = 0
            if ([int]::TryParse($numeroLimpo, [ref]$parsed)) {
                if ($parsed -gt 0 -and $parsed -le $servicosFiltrados.Count) {
                    $indices += ($parsed - 1)
                }
            }
        }
        $selecionados = $indices | ForEach-Object { $servicosFiltrados[$_] }
    }

    if (-not $selecionados -or $selecionados.Count -eq 0) {
        Write-Warning "Nenhum serviço válido selecionado."
        return
    }

    foreach ($servico in $selecionados) {
        $caminhoExe = $servico.PathName -replace '"', ''
        $diretorio = Split-Path $caminhoExe
        $configPath = Join-Path $diretorio "Benner.Tecnologia.Tasks.Worker.exe.config"

        if (-not (Test-Path $configPath)) {
            Write-Warning "⚠️ Arquivo não encontrado: $configPath"
            continue
        }

        try {
            [xml]$xml = Get-Content $configPath
            $node = $xml.configuration.appSettings.add | Where-Object { $_.key -eq "UseCOMFree" }

            if ($node) {
                $valorAntigo = $node.value.ToLower()

                if ($valorAntigo -ne $NovoValor) {
                    $node.value = "$NovoValor"
                    $xml.Save($configPath)
                    Write-Host "✅ $($servico.DisplayName): $valorAntigo → $NovoValor"

                    Registrar-Log -Servico $servico.DisplayName -Parametro "UseCOMFree" -ValorAntigo $valorAntigo -ValorNovo $NovoValor -ArquivoConfig $configPath

                    Write-Host "🔄 Reiniciando o serviço '$($servico.DisplayName)'..."
                    try {
                        Stop-Service -Name $servico.Name -Force -ErrorAction Stop
                        Start-Service -Name $servico.Name -ErrorAction Stop
                        Write-Host "✅ Serviço '$($servico.DisplayName)' reiniciado com sucesso." -ForegroundColor Green
                    } catch {
                        Write-Warning "❌ Erro ao reiniciar o serviço '$($servico.Name)': $_"
                    }
                } else {
                    Write-Host "ℹ️ $($servico.DisplayName) já está com o valor desejado ($NovoValor)" -ForegroundColor Yellow
                }
            } else {
                Write-Warning "⚠️ Parâmetro 'UseCOMFree' não encontrado."
            }

        } catch {
            Write-Warning "❌ Erro ao processar ${configPath}: $_"
        }
    }

    Write-Host "`n✔️ Atualização concluída." -ForegroundColor Green
}

# Função para atualizar o parâmetro SISTEMA no worker ===============================================================================================

Function Atualizar-Sistema {
    # Lista os serviços disponíveis
    $servicos = Get-WmiObject Win32_Service | Where-Object { $_.DisplayName -like "Benner Tasks Worker*" }

    if (-not $servicos.Count) {
        Write-Host "❌ Nenhum serviço encontrado com nome 'Benner Tasks Worker*'."
        return
    }

    Write-Host "`n📋 Serviços encontrados:`n"
    for ($i = 0; $i -lt $servicos.Count; $i++) {
        Write-Host "$($i + 1)) $($servicos[$i].DisplayName)"
    }

    do {
        $escolha = Read-Host "`nDigite o número do serviço que deseja alterar"
        $indice = ($escolha -as [int]) - 1
        $valido = $indice -ge 0 -and $indice -lt $servicos.Count
        if (-not $valido) { Write-Warning "Número inválido. Tente novamente." }
    } while (-not $valido)

    $servico = $servicos[$indice]

    # Solicita novo valor
    do {
        $NovoValor = Read-Host "Informe o novo valor para o parâmetro 'Sistema'"
        $valido = -not [string]::IsNullOrWhiteSpace($NovoValor)
        if (-not $valido) { Write-Warning "❌ Valor não pode ser vazio." }
    } while (-not $valido)

    $caminhoExe = $servico.PathName -replace '"', ''
    $diretorio = Split-Path $caminhoExe
    $configPath = Join-Path $diretorio "Benner.Tecnologia.Tasks.Worker.exe.config"

    if (-not (Test-Path $configPath)) {
        Write-Warning "⚠️ Arquivo não encontrado: $configPath"
        return
    }

    try {
        [xml]$xml = Get-Content $configPath
        $node = $xml.configuration.appSettings.add | Where-Object { $_.key -eq "Sistema" }

        if ($node) {
            $valorAntigo = $node.value

            if ($valorAntigo -ne $NovoValor) {
                $node.value = "$NovoValor"
                $xml.Save($configPath)
                Write-Host "✅ $($servico.DisplayName): $valorAntigo → $NovoValor"

                # Log da alteração
                Registrar-Log -Servico $servico.DisplayName -Parametro "Sistema" -ValorAntigo $valorAntigo -ValorNovo $NovoValor -ArquivoConfig $configPath

                # Reinicia o serviço
                Write-Host "🔄 Reiniciando o serviço '$($servico.DisplayName)'..."
                try {
                    Stop-Service -Name $servico.Name -Force -ErrorAction Stop
                    Start-Service -Name $servico.Name -ErrorAction Stop
                    Write-Host "✅ Serviço '$($servico.DisplayName)' reiniciado com sucesso." -ForegroundColor Green
                } catch {
                    Write-Warning "❌ Erro ao reiniciar o serviço '$($servico.Name)': $_"
                }
            } else {
                Write-Host "ℹ️ $($servico.DisplayName) já possui o valor desejado ($NovoValor)" -ForegroundColor Yellow
            }
        } else {
            Write-Warning "⚠️ Parâmetro 'Sistema' não encontrado em: $configPath"
        }

    } catch {
        Write-Warning "❌ Erro ao processar ${configPath}: $_"
    }

    Write-Host "`n✔️ Atualização concluída." -ForegroundColor Green
}

# Função atualizar pool dinamica Worker =============================================================================================================

Function Atualizar-PoolDinamica {
        # Pergunta sobre tipo de serviço (Agendamento, BTL ou Todos)
        Write-Host "`nDeseja alterar os serviços de:"
        Write-Host "1 - Apenas Agendamento"
        Write-Host "2 - Apenas BTL"
        Write-Host "3 - Todos"
        $tipoFiltro = Read-Host "Digite 1, 2 ou 3"

        do {
            $NovoValor = Read-Host "Informe o novo valor para 'PoolDinamica' (true ou false)"
            $NovoValor = $NovoValor.Trim().ToLower()
            $valido = $NovoValor -in @("true", "false")
            if (-not $valido) { Write-Warning "Digite apenas 'true' ou 'false'." }
        } while (-not $valido)
    
        # Filtra com base na Fila configurada no .config
        $servicosFiltrados = Obter-ServicosPorFila -tipoFiltro $tipoFiltro
        if (-not $servicosFiltrados -or $servicosFiltrados.Count -eq 0) { return }
    
        $selecionados = Selecionar-ServicosInterativamente -ListaDeServicos $servicos
        if (-not $selecionados.Count) { return } 

    foreach ($servico in $selecionados) {
        $caminhoExe = $servico.PathName -replace '"', ''
        $diretorio = Split-Path $caminhoExe
        $configPath = Join-Path $diretorio "Benner.Tecnologia.Tasks.Worker.exe.config"

        if (-not (Test-Path $configPath)) {
            Write-Warning "⚠️ Arquivo não encontrado: $configPath"
            continue
        }

        try {
            [xml]$xml = Get-Content $configPath
            $node = $xml.configuration.appSettings.add | Where-Object { $_.key -eq "PoolDinamica" }

            if ($node) {
                $valorAntigo = $node.value.ToLower()

                if ($valorAntigo -ne $NovoValor) {
                    $node.value = "$NovoValor"
                    $xml.Save($configPath)
                    Write-Host "✅ $($servico.DisplayName): $valorAntigo → $NovoValor"

                    # Log da alteração
                    Registrar-Log -Servico $servico.DisplayName -Parametro "PoolDinamica" -ValorAntigo $valorAntigo -ValorNovo $NovoValor -ArquivoConfig $configPath

                    # Reinicia o serviço forçadamente
                    Write-Host "🔄 Reiniciando o serviço '$($servico.DisplayName)'..."
                    try {
                        Stop-Service -Name $servico.Name -Force -ErrorAction Stop
                        Start-Service -Name $servico.Name -ErrorAction Stop
                        Write-Host "✅ Serviço '$($servico.DisplayName)' reiniciado com sucesso." -ForegroundColor Green
                    } catch {
                        Write-Warning "❌ Erro ao reiniciar o serviço '$($servico.DisplayName)': $_"
                    }
                } else {
                    Write-Host "ℹ️ $($servico.DisplayName) já está com o valor desejado ($NovoValor)" -ForegroundColor Yellow
                }
            } else {
                Write-Warning "⚠️ Parâmetro 'PoolDinamica' não encontrado."
            }

        } catch {
            Write-Warning "❌ Erro ao processar ${configPath}: $_"
        }
    }

    Write-Host "`n✔️ Atualização concluída." -ForegroundColor Green
}

# Função para alterar a fila do worker ==============================================================================================================

Function Atualizar-Fila {

    # Solicita novo valor da Fila (com validação)
    do {
        $NovoValor = Read-Host "Informe o novo valor para o parâmetro 'Fila' (Z_BTLREQUISICOES ou Z_AGENDAMENTOREQUISICOES)"
        $NovoValor = $NovoValor.Trim().ToUpper()
        $valido = $NovoValor -in @("Z_BTLREQUISICOES", "Z_AGENDAMENTOREQUISICOES")
        if (-not $valido) { Write-Warning "❌ Valor inválido. Somente 'Z_BTLREQUISICOES' ou 'Z_AGENDAMENTOREQUISICOES' são permitidos." }
    } while (-not $valido)

    # Lista os serviços disponíveis
    $servicos = Get-WmiObject Win32_Service | Where-Object { $_.DisplayName -like "Benner Tasks Worker*" }

    if (-not $servicos.Count) {
        Write-Host "❌ Nenhum serviço encontrado com nome 'Benner Tasks Worker*'."
        return
    }

    Write-Host "`n📋 Serviços encontrados:`n"
    for ($i = 0; $i -lt $servicos.Count; $i++) {
        Write-Host "$($i + 1)) $($servicos[$i].DisplayName)"
    }

    # Seleção dos serviços
    $selecionados = @()
    $opcao = Read-Host "`nDigite os números dos serviços separados por vírgula (ou 'todos' para alterar todos)"

    if ($opcao.Trim().ToLower() -eq "todos") {
        $selecionados = $servicos
    } else {
        $indices = $opcao -split "," | ForEach-Object { ($_ -as [int]) - 1 }
        $selecionados = $indices | Where-Object { $_ -ge 0 -and $_ -lt $servicos.Count } | ForEach-Object { $servicos[$_] }
    }

    if (-not $selecionados.Count) {
        Write-Warning "Nenhum serviço válido selecionado."
        return
    }

    foreach ($servico in $selecionados) {
        $caminhoExe = $servico.PathName -replace '"', ''
        $diretorio = Split-Path $caminhoExe
        $configPath = Join-Path $diretorio "Benner.Tecnologia.Tasks.Worker.exe.config"

        if (-not (Test-Path $configPath)) {
            Write-Warning "⚠️ Arquivo não encontrado: $configPath"
            continue
        }

        try {
            [xml]$xml = Get-Content $configPath
            $node = $xml.configuration.appSettings.add | Where-Object { $_.key -eq "Fila" }

            if ($node) {
                $valorAntigo = $node.value

                if ($valorAntigo -ne $NovoValor) {
                    $node.value = "$NovoValor"
                    $xml.Save($configPath)
                    Write-Host "✅ $($servico.DisplayName): $valorAntigo → $NovoValor"

                    # Log da alteração
                    Registrar-Log -Servico $servico.DisplayName -Parametro "Fila" -ValorAntigo $valorAntigo -ValorNovo $NovoValor -ArquivoConfig $configPath

                    # Reinicia o serviço
                    Write-Host "🔄 Reiniciando o serviço '$($servico.DisplayName)'..."
                    try {
                        Stop-Service -Name $servico.Name -Force -ErrorAction Stop
                        Start-Service -Name $servico.Name -ErrorAction Stop
                        Write-Host "✅ Serviço '$($servico.DisplayName)' reiniciado com sucesso." -ForegroundColor Green
                    } catch {
                        Write-Warning "❌ Erro ao reiniciar o serviço '$($servico.Name)': $_"
                    }
                } else {
                    Write-Host "ℹ️ $($servico.DisplayName) já possui o valor desejado ($NovoValor)" -ForegroundColor Yellow
                }
            } else {
                Write-Warning "⚠️ Parâmetro 'Fila' não encontrado em: $configPath"
            }

        } catch {
            Write-Warning "❌ Erro ao processar ${configPath}: $_"
        }
    }

    Write-Host "`n✔️ Atualização concluída." -ForegroundColor Green
}

# Atualizar parâmetro SERVIDOR no worker ============================================================================================================

Function Atualizar-Servidor {

    # Solicita novo valor
    do {
        $NovoValor = Read-Host "Informe o novo valor para o parâmetro 'Servidor'"
        $valido = -not [string]::IsNullOrWhiteSpace($NovoValor)
        if (-not $valido) { Write-Warning "❌ Valor não pode ser vazio." }
    } while (-not $valido)

    # Lista os serviços disponíveis
    $servicos = Get-WmiObject Win32_Service | Where-Object { $_.DisplayName -like "Benner Tasks Worker*" }

    if (-not $servicos.Count) {
        Write-Host "❌ Nenhum serviço encontrado com nome 'Benner Tasks Worker*'."
        return
    }

    Write-Host "`n📋 Serviços encontrados:`n"
    for ($i = 0; $i -lt $servicos.Count; $i++) {
        Write-Host "$($i + 1)) $($servicos[$i].DisplayName)"
    }

    $selecionados = @()
    $opcao = Read-Host "`nDigite os números dos serviços separados por vírgula (ou 'todos' para alterar todos)"

    if ($opcao.Trim().ToLower() -eq "todos") {
        $selecionados = $servicos
    } else {
        $indices = $opcao -split "," | ForEach-Object { ($_ -as [int]) - 1 }
        $selecionados = $indices | Where-Object { $_ -ge 0 -and $_ -lt $servicos.Count } | ForEach-Object { $servicos[$_] }
    }

    if (-not $selecionados.Count) {
        Write-Warning "Nenhum serviço válido selecionado."
        return
    }

    foreach ($servico in $selecionados) {
        $caminhoExe = $servico.PathName -replace '"', ''
        $diretorio = Split-Path $caminhoExe
        $configPath = Join-Path $diretorio "Benner.Tecnologia.Tasks.Worker.exe.config"

        if (-not (Test-Path $configPath)) {
            Write-Warning "⚠️ Arquivo não encontrado: $configPath"
            continue
        }

        try {
            [xml]$xml = Get-Content $configPath
            $node = $xml.configuration.appSettings.add | Where-Object { $_.key -eq "Servidor" }

            if ($node) {
                $valorAntigo = $node.value

                if ($valorAntigo -ne $NovoValor) {
                    $node.value = "$NovoValor"
                    $xml.Save($configPath)
                    Write-Host "✅ $($servico.DisplayName): $valorAntigo → $NovoValor"

                    # Log da alteração
                    Registrar-Log -Servico $servico.DisplayName -Parametro "Servidor" -ValorAntigo $valorAntigo -ValorNovo $NovoValor -ArquivoConfig $configPath

                    # Reinicia o serviço
                    Write-Host "🔄 Reiniciando o serviço '$($servico.DisplayName)'..."
                    try {
                        Stop-Service -Name $servico.Name -Force -ErrorAction Stop
                        Start-Service -Name $servico.Name -ErrorAction Stop
                        Write-Host "✅ Serviço '$($servico.DisplayName)' reiniciado com sucesso." -ForegroundColor Green
                    } catch {
                        Write-Warning "❌ Erro ao reiniciar o serviço '$($servico.Name)': $_"
                    }
                } else {
                    Write-Host "ℹ️ $($servico.DisplayName) já possui o valor desejado ($NovoValor)" -ForegroundColor Yellow
                }
            } else {
                Write-Warning "⚠️ Parâmetro 'Servidor' não encontrado em: $configPath"
            }

        } catch {
            Write-Warning "❌ Erro ao processar ${configPath}: $_"
        }
    }

    Write-Host "`n✔️ Atualização concluída." -ForegroundColor Green
}

# Função para remover, incluir e atualizar o parâmetro MultiplicadorCPU =============================================================================

Function Atualizar-MultiplicadorCPU {

    # Pergunta sobre tipo de serviço (Agendamento, BTL ou Todos)
    Write-Host "`nDeseja alterar os serviços de:"
    Write-Host "1 - Apenas Agendamento"
    Write-Host "2 - Apenas BTL"
    Write-Host "3 - Todos"
    $tipoFiltro = Read-Host "Digite 1, 2 ou 3"

    # Filtrar por Fila no .config
    $servicosFiltrados = Obter-ServicosPorFila -tipoFiltro $tipoFiltro
    if (-not $servicosFiltrados -or $servicosFiltrados.Count -eq 0) { return }

    do {
        $NovoValor = Read-Host "Informe o novo valor para o parâmetro 'MultiplicadorCPU' (somente números ou digite 'remover' para remover o parâmetro)"
        $valido = ($NovoValor -match '^[0-9]+$') -or ($NovoValor.Trim().ToLower() -eq "remover")
        if (-not $valido) { Write-Warning "❌ Valor inválido. Digite apenas números ou 'remover'." }
    } while (-not $valido)

    $remover = $NovoValor.Trim().ToLower() -eq "remover"

    # Seleção interativa
    $selecionados = Selecionar-ServicosInterativamente -ListaDeServicos $servicosFiltrados
    if (-not $selecionados.Count) { return }

    if (-not $selecionados -or $selecionados.Count -eq 0) {
        Write-Warning "Nenhum serviço válido selecionado."
        return
    }

    foreach ($servico in $selecionados) {
        $caminhoExe = $servico.PathName -replace '"', ''
        $diretorio = Split-Path $caminhoExe
        $configPath = Join-Path $diretorio "Benner.Tecnologia.Tasks.Worker.exe.config"

        if (-not (Test-Path $configPath)) {
            Write-Warning "⚠️ Arquivo não encontrado: $configPath"
            continue
        }

        try {
            [xml]$xml = Get-Content $configPath
            $appSettings = $xml.configuration.appSettings
            $node = $appSettings.add | Where-Object { $_.key -eq "MultiplicadorCPU" }

            if ($remover) {
                if ($node) {
                    $valorAntigo = $node.value
                    $appSettings.RemoveChild($node) | Out-Null
                    $xml.Save($configPath)
                    Write-Host "❌ $($servico.DisplayName): Parâmetro 'MultiplicadorCPU' removido."
                    Registrar-Log -Servico $servico.DisplayName -Parametro "MultiplicadorCPU" -ValorAntigo $valorAntigo -ValorNovo "REMOVIDO" -ArquivoConfig $configPath

                    Write-Host "🔄 Reiniciando o serviço '$($servico.DisplayName)'..."
                    try {
                        Stop-Service -Name $servico.Name -Force -ErrorAction Stop
                        Start-Service -Name $servico.Name -ErrorAction Stop
                        Write-Host "✅ Serviço '$($servico.DisplayName)' reiniciado com sucesso." -ForegroundColor Green
                    } catch {
                        Write-Warning "❌ Erro ao reiniciar o serviço '$($servico.Name)': $_"
                    }
                } else {
                    Write-Host "ℹ️ $($servico.DisplayName) não possui o parâmetro 'MultiplicadorCPU' para remoção." -ForegroundColor Yellow
                }
            } else {
                if ($node) {
                    $valorAntigo = $node.value
                    if ($valorAntigo -ne $NovoValor) {
                        $node.value = "$NovoValor"
                        $xml.Save($configPath)
                        Write-Host "✅ $($servico.DisplayName): $valorAntigo → $NovoValor"
                        Registrar-Log -Servico $servico.DisplayName -Parametro "MultiplicadorCPU" -ValorAntigo $valorAntigo -ValorNovo $NovoValor -ArquivoConfig $configPath

                        Write-Host "🔄 Reiniciando o serviço '$($servico.DisplayName)'..."
                        try {
                            Stop-Service -Name $servico.Name -Force -ErrorAction Stop
                            Start-Service -Name $servico.Name -ErrorAction Stop
                            Write-Host "✅ Serviço '$($servico.DisplayName)' reiniciado com sucesso." -ForegroundColor Green
                        } catch {
                            Write-Warning "❌ Erro ao reiniciar o serviço '$($servico.Name)': $_"
                        }
                    } else {
                        Write-Host "ℹ️ $($servico.DisplayName) já possui o valor desejado ($NovoValor)" -ForegroundColor Yellow
                    }
                } else {
                    $newNode = $xml.CreateElement("add")
                    $newNode.SetAttribute("key", "MultiplicadorCPU")
                    $newNode.SetAttribute("value", "$NovoValor")
                    $appSettings.AppendChild($newNode) | Out-Null
                    $xml.Save($configPath)
                    Write-Host "➕ $($servico.DisplayName): Parâmetro 'MultiplicadorCPU' adicionado com valor $NovoValor"
                    Registrar-Log -Servico $servico.DisplayName -Parametro "MultiplicadorCPU" -ValorAntigo "N/A" -ValorNovo $NovoValor -ArquivoConfig $configPath

                    Write-Host "🔄 Reiniciando o serviço '$($servico.DisplayName)'..."
                    try {
                        Stop-Service -Name $servico.Name -Force -ErrorAction Stop
                        Start-Service -Name $servico.Name -ErrorAction Stop
                        Write-Host "✅ Serviço '$($servico.DisplayName)' reiniciado com sucesso." -ForegroundColor Green
                    } catch {
                        Write-Warning "❌ Erro ao reiniciar o serviço '$($servico.Name)': $_"
                    }
                }
            }

        } catch {
            Write-Warning "❌ Erro ao processar ${configPath}: $_"
        }
    }

    Write-Host "`n✔️ Atualização concluída." -ForegroundColor Green
}

Function Set-UserForBennerWorkers {
    # Lista serviços que começam com "Benner Tasks Worker*"
    $servicos = Get-WmiObject Win32_Service | Where-Object { $_.DisplayName -like "Benner Tasks Worker*" }

    if (-not $servicos) {
        Write-Warning "Nenhum serviço 'Benner Tasks Worker*' encontrado."
        return
    }

    Write-Host "🔎 Serviços encontrados:" -ForegroundColor Cyan
    $i = 1
    $servicos | ForEach-Object {
        Write-Host "$i. $($_.DisplayName)"
        $i++
    }

    Write-Host ""
    Write-Host "Digite o número do serviço que deseja configurar," -NoNewline
    Write-Host " ou pressione Enter para configurar TODOS." -ForegroundColor Yellow
    $input = Read-Host "Seleção"

    if ([string]::IsNullOrWhiteSpace($input)) {
        $alvo = $servicos
    } elseif ($input -as [int] -and $input -ge 1 -and $input -le $servicos.Count) {
        $alvo = @($servicos[$input - 1])
    } else {
        Write-Warning "❌ Entrada inválida."
        return
    }

    # Solicita credenciais
    $defaultUser = "$env:COMPUTERNAME\bennerservice"
    $cred = Get-Credential -UserName $defaultUser -Message "Informe o usuário e senha que os serviços devem utilizar"

    $usuario = $cred.UserName
    $senha = $cred.GetNetworkCredential().Password

    Write-Host "`n➡️ Conta informada: $usuario" -ForegroundColor Cyan

    foreach ($svc in $alvo) {
        $nome = $svc.Name
        Write-Host "🔧 Configurando logon do serviço '$($svc.DisplayName)'..."

        $resultado = sc.exe config "$nome" obj= "$usuario" password= "$senha"

        if ($LASTEXITCODE -eq 0) {
            Write-Host "✅ Serviço '$nome' configurado para rodar como '$usuario'." -ForegroundColor Green
        } else {
            Write-Warning "❌ Falha ao configurar o serviço '$nome'."
            Write-Output $resultado
        }
    }

    Write-Host "`n✔️ Configuração finalizada." -ForegroundColor Cyan
}

Function Uninstall-BennerWorkers {
    $servicos = Get-WmiObject Win32_Service | Where-Object { $_.DisplayName -like "Benner Tasks Worker*" }

    if (-not $servicos) {
        Write-Warning "Nenhum serviço 'Benner Tasks Worker*' encontrado."
        return
    }

    Write-Host "🔎 Serviços encontrados:" -ForegroundColor Cyan
    $i = 1
    $servicos | ForEach-Object {
        Write-Host ("{0}. {1} (Usuário: {2}, Startup: {3})" -f $i, $_.DisplayName, $_.StartName, $_.StartMode)
        $i++
    }

    Write-Host ""
    Write-Host "Digite os números dos serviços a desinstalar (ex: 1,3), ou digite 'todos' para remover todos." -ForegroundColor Yellow
    $input = Read-Host "Seleção"

    if ($input.Trim().ToLower() -eq "todos") {
        $alvo = $servicos
    }
    elseif ([string]::IsNullOrWhiteSpace($input)) {
        Write-Warning "❌ Você deve digitar 'todos' ou uma lista de serviços (ex: 1,3)"
        return
    }
    else {
        $indices = $input -split "," | ForEach-Object { ($_ -as [int]) - 1 }
        $alvo = @()
        foreach ($i in $indices) {
            if ($i -ge 0 -and $i -lt $servicos.Count) {
                $alvo += $servicos[$i]
            } else {
                Write-Warning "⚠️ Número inválido: $($i+1)"
            }
        }
        if (-not $alvo) {
            Write-Warning "❌ Nenhum serviço válido selecionado."
            return
        }
    }

    foreach ($svc in $alvo) {
        $nomeServico = $svc.DisplayName
        $path = $svc.PathName -replace '"', ''
        $pasta = Split-Path $path -Parent
        $exe = Join-Path $pasta "Benner.Tecnologia.Tasks.Worker.exe"

        if (Test-Path $exe) {
            Write-Host "🧨 Desinstalando serviço: $nomeServico" -ForegroundColor Yellow
            Push-Location $pasta
            & "$exe" uninstall
            Pop-Location
            Write-Host "✅ Serviço '$nomeServico' desinstalado." -ForegroundColor Green
        } else {
            Write-Warning "❌ Executável não encontrado para '$nomeServico' em '$pasta'"
        }
    }

    Write-Host "`n✔️ Desinstalação concluída." -ForegroundColor Cyan
}

Function Reiniciar-Workers {
    $servicos = Get-WmiObject Win32_Service | Where-Object { $_.DisplayName -like "Benner Tasks Worker*" }

    if (-not $servicos.Count) {
        Write-Host "❌ Nenhum serviço encontrado com nome 'Benner Tasks Worker*'." -ForegroundColor Yellow
        return
    }

    Write-Host "`n📋 Serviços encontrados:`n"
    for ($i = 0; $i -lt $servicos.Count; $i++) {
        Write-Host "$($i + 1)) $($servicos[$i].DisplayName)"
    }

    $selecionados = @()
    $opcao = Read-Host "`nDigite os números dos serviços separados por vírgula (ou 'todos' para reiniciar todos)"

    if ($opcao.Trim().ToLower() -eq "todos") {
        $selecionados = $servicos
    } else {
        $indices = @()
        $opcao.Split(",") | ForEach-Object {
            $numeroLimpo = $_.Trim()
            $parsed = 0
            if ([int]::TryParse($numeroLimpo, [ref]$parsed)) {
                if ($parsed -gt 0 -and $parsed -le $servicos.Count) {
                    $indices += ($parsed - 1)
                }
            }
        }
        $selecionados = $indices | ForEach-Object { $servicos[$_] }
    }

    if (-not $selecionados -or $selecionados.Count -eq 0) {
        Write-Warning "Nenhum serviço válido selecionado."
        return
    }

    foreach ($servico in $selecionados) {
        Write-Host "`n🔄 Reiniciando o serviço '$($servico.DisplayName)'..."
        try {
            Stop-Service -Name $servico.Name -Force -ErrorAction Stop
            Start-Service -Name $servico.Name -ErrorAction Stop
            Write-Host "✅ Serviço '$($servico.DisplayName)' reiniciado com sucesso." -ForegroundColor Green
        } catch {
            Write-Warning "❌ Erro ao reiniciar o serviço '$($servico.Name)': $_"
        }
    }

    Write-Host "`n✔️ Reprocesso de reinicialização concluído." -ForegroundColor Green
}

function ExportarConfigWorker {

    Write-Host "`nDeseja alterar os serviços de:"
    Write-Host "1 - Apenas Agendamento"
    Write-Host "2 - Apenas BTL"
    Write-Host "3 - Todos"
    $tipoFiltro = Read-Host "Digite 1, 2 ou 3"

    $servicosFiltrados = Obter-ServicosPorFila -tipoFiltro $tipoFiltro

    if (-not $servicosFiltrados -or $servicosFiltrados.Count -eq 0) {
        return
    }

    # Início da declaração
    $saida = "`$workers = @(`n"

    $contador = 1

    foreach ($svc in $servicosFiltrados) {
        $exePath = $svc.PathName -replace '"', ''  # Remove aspas
        $installDir = Split-Path $exePath -Parent
        $configPath = Join-Path $installDir "Benner.Tecnologia.Tasks.Worker.exe.config"
        $pasta = Split-Path $installDir -Leaf

        if (Test-Path $configPath) {
            [xml]$xml = Get-Content $configPath

            # Extrai os valores do XML
            $settings = @{ }
            foreach ($add in $xml.configuration.appSettings.add) {
                $settings[$add.key] = $add.value
            }

            # Ajusta NomeServico
            $nomeServico = if ($settings.ContainsKey("NomeServico") -and $settings["NomeServico"].Trim() -ne "") {
                $settings["NomeServico"]
            } else {
                $pasta
            }

            # Mostra ao usuário antes de exportar
            Write-Host "` $contador - Serviço: $($svc.DisplayName) " -ForegroundColor Green
            Write-Host "Pasta: $pasta | Sistema: $($settings["Sistema"]) | Servidor: $($settings["Servidor"]) | NumeroDeProviders: $($settings["NumeroDeProviders"]) | Usuario: $($settings["Usuario"]) | Fila: $($settings["Fila"]) | UseCOMFree: $($settings["UseCOMFree"]) | LoggingServerActive: $($settings["LoggingServerActive"]) | LoggingServerAddress: $($settings["LoggingServerAddress"]) | NomeServico: $nomeServico | PoolDinamica: $($settings["PoolDinamica"]) | PrimaryMessageFilterSQL: $($settings["PrimaryMessageFilterSQL"]) | SecondaryMessageFilterSQL: $($settings["SecondaryMessageFilterSQL"]) | MultiplicadorCPU: $($settings["MultiplicadorCPU"])"
        
            <#Write-Host "`n📋 Serviços encontrados:`n"
            for ($i = 0; $i -lt $servicosFiltrados.Count; $i++) {
                Write-Host "$($i + 1)) $($servicosFiltrados[$i].DisplayName)"
            } #>

            # Constrói a entrada exportável
            $saida += "    @{ Pasta = `"$pasta`";"
            $saida += " Sistema = `"$($settings["Sistema"])`";"
            $saida += " Servidor = `"$($settings["Servidor"])`";"
            $saida += " NumeroDeProviders = $($settings["NumeroDeProviders"] -as [int]);"
            $saida += " Usuario = `"$($settings["Usuario"])`";"
            $saida += " Fila = `"$($settings["Fila"])`";"
            $saida += " UseCOMFree = `"$($settings["UseCOMFree"])`";"
            $saida += " LoggingServerActive = `"$($settings["LoggingServerActive"])`";"
            $saida += " LoggingServerAddress = `"$($settings["LoggingServerAddress"])`";"
            $saida += " NomeServico = `"$nomeServico`";"
            $saida += " PoolDinamica = `"$($settings["PoolDinamica"])`";"

            if ($settings.ContainsKey("PrimaryMessageFilterSQL")) {
                $saida += " PrimaryMessageFilterSQL = `"$($settings["PrimaryMessageFilterSQL"])`";"
            }
            if ($settings.ContainsKey("SecondaryMessageFilterSQL")) {
                $saida += " SecondaryMessageFilterSQL = `"$($settings["SecondaryMessageFilterSQL"])`";"
            }
            if ($settings.ContainsKey("MultiplicadorCPU")) {
                $saida += " MultiplicadorCPU = $($settings["MultiplicadorCPU"] -as [int]);"
            }

            $saida += " ServiceName = `"BennerTasksWorker_$([guid]::NewGuid())`" },`n"
            $contador++
        } else {
            Write-Warning "⚠️ Arquivo de configuração não encontrado para o serviço '$($svc.DisplayName)' em '$installDir'"
        }
    }

    # Remove vírgula final e fecha o array
    $saida = $saida.TrimEnd("`,", "`n") + "`n)"

    # Pergunta onde salvar o arquivo
    $basePath = Read-Host "Informe o diretório onde deseja salvar o arquivo (ex: C:\TEMP)"

    if (!(Test-Path $basePath)) {
        Write-Warning "⚠ O diretório '$basePath' não existe."
        Return
    } else {
        # Exporta o conteúdo
        $arquivoDestino = Join-Path $basePath "export_workers.ps1"
        $saida | Out-File -FilePath $arquivoDestino -Encoding UTF8

        Write-Host "`n✅ Arquivo exportado com sucesso para: $arquivoDestino"
    }
}


# MENU INICIAL
do {
    Write-Host "`n===== MENU DE OPÇÕES ====="
    Write-Host "1) Atualizar parametro NumeroDeProviders"
    Write-Host "2) Atualizar parametro UseCOMFree"
    Write-Host "3) Atualizar parametro Sistema"
    Write-Host "4) Atualizar parametro PoolDinamica"
    Write-Host "5) Atualizar parametro Fila"
    Write-Host "6) Atualizar parametro Servidor"
    Write-Host "7) Atualizar parametro MultiplicadorCPU"
    Write-Host "8) Atualizar usuário de serviço"
    Write-Host "9) Desinstalar serviço worker"
    Write-Host "10) Reiniciar Worker"
    Write-Host "11) Desabilitar e Habilitar Worker"
    Write-Host "12) Exportar configuração worker"
    Write-Host "0) Sair"
    $opcao = Read-Host "Escolha a opção desejada"

    switch ($opcao) {
        "1" { Atualizar-NumeroDeProviders }
        "2" { Atualizar-UseCOMFree }
        "3" { Atualizar-Sistema }
        "4" { Atualizar-PoolDinamica }
        "5" { Atualizar-Fila }
        "6" { Atualizar-Servidor }
        "7" { Atualizar-MultiplicadorCPU }
        "8" { Set-UserForBennerWorkers }
        "9" { Uninstall-BennerWorkers }
        "10" { Reiniciar-Workers }
        "11" { Gerenciar-BennerTasksWorker }
        "12" { ExportarConfigWorker }
        "0" { Write-Host "`nSaindo..." }
        default { Write-Warning "Opção inválida. Tente novamente." }
    }

} while ($opcao -ne "0")
