def obter_cliente_por_chamado(numero_chamado):
    # Função que recebe o número do chamado e retorna o cliente associado a ele
    # Implemente a lógica para mapear números de chamado a clientes aqui
    if numero_chamado.startswith('55'):
        return 'MAPA'
    elif numero_chamado.startswith('22'):
        return 'PCDF'
    elif numero_chamado.startswith('96'):
        return 'AGU'
    elif numero_chamado.startswith('83'):
        return 'ANA'
    elif numero_chamado.startswith('40'):
        return 'ANM'
    elif numero_chamado.startswith('78'):
        return 'MINFRA'
    elif numero_chamado.startswith('15'):
        return 'MMA'
    elif numero_chamado.startswith('81'):
        return 'MMA'
    elif numero_chamado.startswith('081'):
        return 'MMA'
    else:
        return 'Cliente Desconhecido'