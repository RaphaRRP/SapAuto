import tabela, janelasSap, mb25, transacoes

def main():
    janelasSap.janelasSap()
    mb25.rodar_mb25()
    transacoes.rodar_transacoes_sap()

def submain():
    transacoes.exportar_051()
    tabela.alterar_tabelas()



