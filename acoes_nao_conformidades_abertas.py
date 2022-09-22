#Importando bibliotecas
import win32com.client
import datetime
import time
import sys
from dateutil.relativedelta import relativedelta

hoje = datetime.datetime.now()
ano = hoje.year
mes = hoje.month
dia = hoje.day
hora = hoje.hour
minuto = hoje.minute

#Começando uma instância em Excel
xlapp = win32com.client.DispatchEx("Excel.Application")

#Passo opcional: Para debug
xlapp.Visible = 0

#Abrir o workbook na instância mencionada (Excel)
wb = xlapp.workbooks.open("C:\\Users\\carlos.junior\\Desktop\\Dashboards\\Controle acoes de nao conformidades em aberto.lnk")


#Atualizando todas as conexões
wb.RefreshAll()

#Mantém o programa aguardando até a sincronização ser concluída
xlapp.CalculateUntilAsyncQueriesDone()

#Executa uma macro que atualiza o campo "Última atualização"
xlapp.Application.Run("ultimaAtualizacao")
xlapp.Application.Run("impressao")
xlapp.Application.Run("PintarFiltros")

#Diz quais abas da planilha que devem ser impressas
ws_index_list = [1]

#Define o nome do arquivo pdf
name_document = f"NC_AcoesEmAberto_{ano}-{mes}-{dia}"

#Define o caminho do relatório
path_to_pdf = f"N:\\12 - Ferramentas de Busca\\Relatorios Semanais\\{name_document}.pdf"

#Seleciona as planilhas definidas no ws_index_list
wb.worksheets(ws_index_list).Select()

#Exportando em formato pdf
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)

#Limpando a área de impressão
xlapp.Application.Run("ClearPrintArea")

#Remover alertas
xlapp.DisplayAlerts = False

#Salvar e sair
wb.Save()
wb.Close()
xlapp.Quit()

#Abrindo a aplicação do Outlook
outlook = win32com.client.Dispatch("Outlook.Application")

#Criando um novo e-mail
Msg = outlook.CreateItem(0)

#Definindo os destinatários
Msg.To = "carlos.junior@hkm.ind.br;engenharia@hkm.ind.br;pcp@hkm.ind.br;sgq@hkm.ind.br;emerson.carvalho@hkm.ind.br;joao.ganda@hkm.ind.br;lsa.inspecao@gmail.com;saude.ocupacional@hkm.ind.br;nilcelia.ferreira@hkm.ind.br;leonardo.clemente@hkm.ind.br;wilson.silva@hkm.ind.br"
Msg.BCC = "thiago.pereira@hkm.ind.br;octavio.pereira@hkm.ind.br;ronaldo.silva@hkm.ind.br;carlos.cirino@hkm.ind.br"

#Assunto do e-mail
Msg.Subject = f"Ações em aberto {dia}-{mes}-{ano}"

#Corpo do e-mail
Msg.Body = f'''
Bom dia,

Em anexo, a relação das ações em aberto das suas respectivas não conformidades.

Em caso de dúvidas ou sugestões, favor entrar em contato.

Este é um e-mail automático, mas sinta-se livre para respondê-lo.
'''

#Anexando o relatório
Msg.Attachments.Add(path_to_pdf)

#Enviar o e-mail
Msg.Send()