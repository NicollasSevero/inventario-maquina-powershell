Add-Type -AssemblyName System.Windows.Forms

$ArquivoInventario = "geral\TI\inventario\INVENTARIO_SH.csv"

try {
    $NomeComputador = $env:COMPUTERNAME
    $UsuarioLogado = $env:USERNAME

    $IPs = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
    $EnderecoIP = @()
    foreach ($ip in $IPs) {
        foreach ($addr in $ip.IPAddress) {
            if ($addr -notlike "*:*" -and $addr -ne "127.0.0.1") {
                $EnderecoIP += $addr
            }
        }
    }
    $IP = $EnderecoIP -join ", "

    $Marca = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
    $Modelo = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
    $Processador = (Get-CimInstance Win32_Processor).Name
    $MemoriaRAM = [math]::round((Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
    $PlacaDeVideoInfo = Get-CimInstance -ClassName Win32_VideoController | Select-Object -ExpandProperty Name
    $PlacaDeVideo = $PlacaDeVideoInfo -join ", "
    $DiscoC = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
    $EspacoLivreC = [math]::Round($DiscoC.FreeSpace / 1GB, 2)
    $EspacoTotal = [math]::Round($DiscoC.Size / 1GB, 2)
    $SistemaOperacional = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption
    $DataHora = Get-Date -Format "dd/MM/yyyy HH:mm"

    # Ordem correta definida explicitamente
    $Inventario = [PSCustomObject]@{
        "DATA/ HORA"           = $DataHora
        "USUARIO"              = $UsuarioLogado
        "NOME DO COMPUTADOR"   = $NomeComputador
        "MARCA"                = $Marca
        "MODELO"               = $Modelo
        "PROCESSADOR"          = $Processador
        "PLACA DE VIDEO"       = $PlacaDeVideo
        "MEMORIA RAMGB"        = $MemoriaRAM
        "SISTEMA OPERACIONAL"  = $SistemaOperacional
        "ESPAÇO LIVRE"         = $EspacoLivreC
        "ESPAÇO TOTAL"         = $EspacoTotal
        "IP"                   = $IP
    }

    if (Test-Path $ArquivoInventario) {
        $LinhasExistentes = @(Import-Csv -Path $ArquivoInventario -Delimiter ";" | Where-Object { $_."NOME DO COMPUTADOR" -ne $NomeComputador })
        $LinhasAtualizadas = @($LinhasExistentes)
        $LinhasAtualizadas += $Inventario
        $LinhasAtualizadas | Export-Csv -Path $ArquivoInventario -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
        $Inventario | Export-Csv -Path $ArquivoInventario -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    }

    [System.Windows.Forms.MessageBox]::Show("Inventário atualizado com sucesso!", "Sucesso", 'OK', 'Information')
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Erro ao atualizar o inventário: $($_.Exception.Message)", "Erro", 'OK', 'Error')
}
