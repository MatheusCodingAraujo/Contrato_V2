import streamlit as st
from datetime import datetime
from docx import Document
import io
from num2words import num2words


def numero_extenso(numero):
    numero_formatado = f"{numero:,.2f}".replace(
        ",", "X").replace(".", ",").replace("X", ".")
    return numero_formatado


def gerar_contrato(nome_cliente, nacionalidade, civil, profissao, cpf_cliente, rg_cliente,
                   endereco_cliente, cep, pedido, descricao_pedido, endereco_obra, valor_pedido,
                   prazo_pedido, parcelas, descricao_parcela, documento_upload):
    # Obtém a data atual
    data_atual = datetime.now()

    # Cria uma lista com os meses em português
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    # Obtém o dia e o mês
    dia = data_atual.day
    mes = meses[data_atual.month - 1]  # O mês é indexado a partir de 0
    ano = data_atual.year

    # Formata a data no estilo desejado
    data_formatada = f"{dia} de {mes} de {ano}"

    nome_cliente = nome_cliente.strip()
    nacionalidade = nacionalidade.lower()
    civil = civil.lower()
    profissao = profissao.lower()
    cpf_cliente = cpf_cliente.strip()
    rg_cliente = rg_cliente.strip()
    endereco_cliente = endereco_cliente.strip()
    cep = cep.strip()
    pedido = pedido.strip()
    descricao_pedido = descricao_pedido.strip()
    endereco_obra = endereco_obra.strip()

    pedido_extenso = numero_extenso(valor_pedido)
    valor_extenso = num2words(valor_pedido, lang='pt_BR', to='currency')
    valor_pedido = f"{pedido_extenso} ({valor_extenso})"

    prazo_extenso = num2words(prazo_pedido, lang='pt_BR')
    prazo_pedido = f"{prazo_pedido} ({prazo_extenso})"

    parcelas_texto = "\n".join(
        [f"{i + 1}) R$ {numero_extenso(valor)} ({num2words(valor, lang='pt_BR', to='currency')}) - {descricao_parcela[i]}"
         for i, valor in enumerate(parcelas)])

    # Lê o documento enviado pelo usuário
    doc = Document(documento_upload)

    # Dicionário de substituições
    substituicoes = {
        '{{nome_cliente}}': nome_cliente,
        '{{nacionalidade}}': nacionalidade,
        '{{civil}}': civil,
        '{{profissao}}': profissao,
        '{{cpf_cliente}}': cpf_cliente,
        '{{rg_cliente}}': rg_cliente,
        '{{endereco_cliente}}': endereco_cliente,
        '{{cep}}': cep,
        '{{pedido}}': pedido,
        '{{descricao_pedido}}': descricao_pedido,
        '{{endereco_obra}}': endereco_obra,
        '{{valor_pedido}}': str(valor_pedido),
        '{{prazo_pedido}}': str(prazo_pedido),
        '{{data}}': data_formatada,
        # Adicionando a substituição para parcelas
        '{{parcelas}}': parcelas_texto,
    }

    # Percorre todos os parágrafos do documento
    for paragrafo in doc.paragraphs:
        for run in paragrafo.runs:  # Percorre cada "run" (parte do texto)
            for chave, valor in substituicoes.items():
                if chave in run.text:
                    # Substitui a chave pelo valor
                    print(f"Substituindo {chave} por {valor}")
                    run.text = run.text.replace(chave, valor)

                    # Aplica negrito somente para nome do cliente e CPF
                    if chave in ['{{nome_cliente}}', '{{cpf_cliente}}']:
                        run.bold = True
                    elif chave in ['{{data}}', '{{civil}}']:
                        run.bold = False  # Garante que a data não fique em negrito

    print("--------------- FIM DE LOOP-----------------")

    # Salva o documento modificado em um arquivo em memória
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)  # Move o ponteiro para o início do arquivo em memória

    return output


# Interface em Streamlit
st.title("Gerador de Contratos")
st.subheader("Envie seu arquivo Word para rápida personalização.")
st.write("Coloque as TAGS no molde para que o sistema possa substituir, NESSE FORMATO:")

# Campos do formulário

# Coleta de dados do cliente
nome_cliente = st.text_input('Nome do Cliente')
nacionalidade = st.text_input('Nacionalidade')
civil = st.text_input('Estado Civil')
profissao = st.text_input('Profissão')
cpf_cliente = st.text_input('CPF do Cliente')
rg_cliente = st.text_input('RG do Cliente')
endereco_cliente = st.text_input('Endereço do Cliente')
cep = st.text_input('CEP')

# Dados do pedido
pedido = st.text_input('Número do Pedido')
descricao_pedido = st.text_area('Descrição do Pedido')
endereco_obra = st.text_input('Endereço da Obra')
valor_pedido = st.number_input('Valor do Pedido', format='%0.2f')
prazo_pedido = st.number_input(
    'Prazo do Pedido', min_value=0, step=1, format='%d')

# Coleta de parcelas
num_parcelas = st.number_input(
    'Número de Parcelas', min_value=1, step=1, format='%d')
valores_parcelas = []
descricoes_parcelas = []

for i in range(int(num_parcelas)):
    valor_parcela = st.number_input(
        f'Valor da Parcela {i + 1}', format='%0.2f')
    descricao_parcela = st.text_input(f'Descrição da Parcela {i + 1}')
    descricoes_parcelas.append(descricao_parcela)
    valores_parcelas.append(valor_parcela)

# Campo para upload do documento
documento_upload = st.file_uploader(
    "Faça o upload do seu arquivo .docx", type="docx")

# Botão para gerar o contrato
if st.button("Gerar Contrato"):
    if (nome_cliente and nacionalidade and civil and profissao and cpf_cliente and rg_cliente and
        endereco_cliente and cep and pedido and descricao_pedido and endereco_obra and
            valor_pedido and prazo_pedido and documento_upload):

        # Gera o contrato e recebe o arquivo modificado
        arquivo_contrato = gerar_contrato(
            nome_cliente, nacionalidade, civil, profissao, cpf_cliente, rg_cliente, endereco_cliente,
            cep, pedido, descricao_pedido, endereco_obra, valor_pedido, prazo_pedido, valores_parcelas,
            descricoes_parcelas, documento_upload)

        # Exibe um botão para baixar o arquivo gerado
        st.download_button(
            label="Baixar Contrato Modificado",
            data=arquivo_contrato,
            file_name=f'Contrato {nome_cliente}_{pedido}.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        st.error("Por favor, preencha todos os campos e faça o upload do documento.")
