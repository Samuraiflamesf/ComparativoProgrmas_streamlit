import pandas as pd
import streamlit as st
from io import BytesIO


# Função para carregar a planilha
def carregar_planilha(file, skiprows=0):
    try:
        planilha = pd.read_excel(file, skiprows=skiprows)
        return planilha
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")
        return None


# Função para salvar o DataFrame em um buffer de bytes
def to_excel_bytes(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Planilha")
    buffer.seek(0)
    return buffer


def main():
    st.title("Estoque SIGAF X SIMPAS")

    st.write(
        """
        Este aplicativo permite comparar planilhas das bases SIMPAS e SIGAF. 
        Por favor, faça o upload das planilhas necessárias e veja os resultados abaixo.
        """
    )

    # Dividindo a interface em duas colunas
    col1, col2 = st.columns(2)

    with col1:
        st.write("#### Selecione a planilha SIMPAS")
        simpas_file = st.file_uploader(
            "Clique para selecionar a planilha SIMPAS (somente .xlsx)", type=["xlsx"]
        )
        if simpas_file:
            st.write(f"Arquivo SIMPAS carregado: {simpas_file.name}")

    with col2:
        st.write("#### Selecione a planilha SIGAF")
        sigaf_file = st.file_uploader(
            "Clique para selecionar a planilha SIGAF (somente .xlsx)", type=["xlsx"]
        )
        if sigaf_file:
            st.write(f"Arquivo SIGAF carregado: {sigaf_file.name}")

    # Botões para processar as planilhas
    if simpas_file and sigaf_file:
        st.write("Planilhas carregadas com sucesso!")

        # Carregar planilhas
        planilha1 = carregar_planilha(simpas_file)
        planilha2 = carregar_planilha(sigaf_file, skiprows=7)

        if planilha1 is not None and planilha2 is not None:
            # Verificar colunas e processar dados
            if "Código" not in planilha1.columns or "Nome" not in planilha1.columns:
                st.error("Colunas necessárias não encontradas na planilha SIMPAS!")
                return

            planilha1 = planilha1[["Código", "Nome", "Saldo"]]
            planilha1["Código"] = (
                planilha1["Código"]
                .astype(str)
                .str.replace(".", "")
                .str.replace("-", "")
            )
            planilha1["Código"] = planilha1["Código"].str.rstrip()

            if "Código Simpas" not in planilha2.columns:
                st.error("Coluna necessária não encontrada na planilha SIGAF!")
                return

            planilha2 = planilha2.rename(columns={"Código Simpas": "Código"})
            planilha2 = planilha2[
                ["Código", "Medicamento", "Quantidade Encontrada", "Programa Saúde"]
            ]
            planilha2 = (
                planilha2.groupby("Código")["Quantidade Encontrada"].sum().reset_index()
            )

            planilha1["Código"] = planilha1["Código"].astype(str)
            planilha2["Código"] = planilha2["Código"].astype(str).str.rstrip()

            # Mesclando os DataFrames
            app_excel = pd.merge(planilha2, planilha1, on="Código", how="left")
            app_excel = app_excel.rename(
                columns={"Quantidade Encontrada": "Saldo SIGAF"}
            )
            app_excel = app_excel.rename(columns={"Saldo": "Saldo SIMPAS"})
            app_excel["Diferença"] = (
                app_excel["Saldo SIGAF"] - app_excel["Saldo SIMPAS"]
            )

            app_excel = app_excel[
                ["Código", "Nome", "Saldo SIMPAS", "Saldo SIGAF", "Diferença"]
            ]

            # Exibir tabela resultante
            st.write("Resultado da Análise:")
            st.dataframe(app_excel)

            # Botão para download do arquivo
            excel_bytes = to_excel_bytes(app_excel)
            st.download_button(
                label="Baixar Arquivo Resultante",
                data=excel_bytes,
                file_name="resultado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
