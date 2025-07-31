# WorkerManager
Automação do serviço do worker

# Gerenciador de Serviços Benner Tasks Worker

Este script PowerShell fornece um menu interativo para **gerenciar, atualizar e exportar configurações** de serviços Windows que utilizam o executável `Benner.Tecnologia.Tasks.Worker.exe`.

Ele permite automatizar alterações no `.config` dos Workers, reiniciar serviços, atualizar parâmetros, definir usuário de execução, exportar configurações e muito mais — com logs de auditoria para rastreabilidade.

---

## 🚀 Funcionalidades disponíveis

### 🔧 Atualizações em configurações `.config`
- `NumeroDeProviders`
- `UseCOMFree`
- `Sistema`
- `PoolDinamica`
- `Fila`
- `Servidor`
- `MultiplicadorCPU` (adicionar, atualizar ou remover)

### ⚙️ Gerenciamento de serviços
- Alterar usuário do serviço (com suporte a senha segura)
- Reiniciar serviços Benner
- Desinstalar serviços
- Desabilitar/habilitar serviços (se a função `Gerenciar-BennerTasksWorker` estiver implementada)

### 📤 Exportação
- Geração de um arquivo `.ps1` com as configurações atuais dos serviços selecionados
- Filtro por tipo de fila (Agendamento ou BTL)
- Visualização prévia no console
- Exportação para diretório definido pelo usuário

### 📝 Log
- Todas as alterações são registradas no arquivo `log_modificacoes_worker.txt`, salvo na mesma pasta do script.
- Cada linha contém: data/hora, parâmetro alterado, valor antigo, novo valor, nome do serviço e caminho do arquivo `.config`.

---

## 📋 Pré-requisitos

- PowerShell 5.1 ou superior
- Permissão de administrador (para editar `.config`, parar/iniciar serviços e alterar credenciais)
- Política de execução liberada durante a sessão:
  
  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
