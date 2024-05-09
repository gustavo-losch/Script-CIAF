import tabula

tabela = tabula.read_pdf("Extrato Cecilia.pdf", pages="all")[0]
tabela.rename(columns=tabela.iloc[4], inplace = True)
tabela = tabela[['DOCTO:', 'VENCIMENTO:', 'R$ DEVIDO:']]
tabela = tabela.drop(0)
tabela = tabela.drop(1)
tabela = tabela.drop(2)
tabela = tabela.drop(3)
tabela = tabela.drop(4)

print(tabela)