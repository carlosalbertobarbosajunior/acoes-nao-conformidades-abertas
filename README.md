# acoes-nao-conformidades-abertas
 Relatório semanal das ações de não conformidades abertas no sistema

Problemática:
	Gerar um relatório automático capaz de extrair as ações de não conformidades pendentes no ERP da empresa e enviá-lo por e-mail aos interessados.
	
	
View em T-SQL: (view-acoes-em-aberto.sql)<br />
	Responsável por realizar a junção da tabela de ações em aberto com a tabela de não conformidades, além de filtrar as não conformidades já concluídas e corrigidas (implementadas).


Script Python: (acoes_nao_conformidades_abertas.py)<br />
	Através da biblioteca win32com, o script é capaz de:<br />
		-Abrir uma instância em excel;<br />
		-Abrir a planilha conectada previamente com o banco de dados;<br />
		-Atualizar todos os vínculos com o banco de dados;<br />
		-Executar a macro vba-impressao-dinamica.bas;<br />
		-Exportar o arquivo em formato PDF;<br />
		-Salvar e fechar a instância do Excel;<br />
		-Abrir uma instância outlook e enviar o arquivo para a lista de e-mails pré-definida.
		

Script VBA: (vba-impressao-dinamica.bas)<br />
	Identifica o comprimento da tabela recém sincronizada com o banco de dados e define a área de impressão de forma dinâmica

