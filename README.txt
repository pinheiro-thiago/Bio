PORTAL STREAMLIT — UNIDADE DE BIOCOMBUSTÍVEIS DE ITAIPU

Conteúdo:
- Página inicial institucional.
- Biblioteca de vídeos com os três links fornecidos.
- Dashboard operacional baseado no arquivo Operação.xlsx, somente leitura.
- Filtros por período e ponto de medição, gráficos e exportação CSV.
- Logo convertido do arquivo PDF enviado.

REQUISITOS
- Recomendado: Python 3.12 de 64 bits.
- Windows: dê duplo clique em Iniciar_Portal.bat.
- Alternativa manual:
    python -m venv .venv
    .venv\Scripts\activate
    python -m pip install -r requirements.txt
    python -m streamlit run app.py

ARQUIVO EXCEL
- O arquivo Operação.xlsx está incluído na pasta e deve permanecer ao lado de app.py.
- A aplicação abre o arquivo em modo somente leitura por meio de openpyxl. Não utiliza Excel/COM nem escreve alterações na planilha.
- Abas esperadas: Operação_CSTR e Composição_Biogás.
- O cache é de 60 segundos; use o botão “Atualizar dados do Excel” para forçar nova leitura.
- Para que alterações feitas por outra pessoa apareçam, o arquivo Excel deve ser atualizado/substituído no local onde o aplicativo está rodando.

VÍDEOS
- Institucional: https://youtu.be/qfgAGSnDA4U
- Segurança: https://youtu.be/RVSyaAXuDuw
- Diário de obra: https://youtu.be/8wWsIYfZUf8

ACESSO POR OUTROS COMPUTADORES
Executar localmente disponibiliza o portal apenas no computador anfitrião por padrão. Para a equipe acessar em diferentes computadores, execute em servidor/VM institucional sempre ligado e configure a rede/firewall com a equipe de TI. Antes de hospedar externamente, confirme as políticas da instituição para dados operacionais e vídeos.

SEGURANÇA
O portal não altera a planilha original. O arquivo Excel está empacotado junto com o código; avalie se ele pode ser distribuído dessa forma antes de compartilhar o ZIP.
