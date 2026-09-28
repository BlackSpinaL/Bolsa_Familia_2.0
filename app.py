"""
Calculadora de Frequência Escolar - Bolsa Família (2026)
-------------------------------------------------------
Nova lógica de cálculo: frequência por DIAS (não por aulas/disciplinas).

Entrada: mês e faltas (em dias).
Calendário fornece: dias letivos de cada mês.
Fórmula: % Frequência = 1 − (faltas ÷ dias letivos)

A estrutura de menus foi mantida igual à versão anterior (4 abas, com as
subtabs "Dias letivos por mês", "Turmas regulares" e "Turmas com contraturno"
dentro de Calendário). As duas últimas ficam apenas como referência — não
entram mais no cálculo.
"""

import json
import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# Configuração da página
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Frequência - Bolsa Família",
    page_icon="📋",
    layout="wide",
)

CONFIG_PATH = Path(__file__).parent / "calendario_2026.json"
LIMITE_FREQUENCIA = 0.75

# Mantidos só para exibição/referência na subtab "Turmas com contraturno"
DIAS_SEMANA = ["segunda", "terca", "quarta", "quinta", "sexta"]
DIAS_SEMANA_LABEL = {
    "segunda": "Segunda",
    "terca": "Terça",
    "quarta": "Quarta",
    "quinta": "Quinta",
    "sexta": "Sexta",
}


# ----------------------------------------------------------------------------
# Calendário (carregar / salvar)
# ----------------------------------------------------------------------------
def carregar_calendario_do_disco(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def garantir_calendario_na_sessao():
    if "calendario" not in st.session_state:
        st.session_state["calendario"] = carregar_calendario_do_disco(CONFIG_PATH)


def salvar_calendario_no_disco():
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(st.session_state["calendario"], f, ensure_ascii=False, indent=2)


def lista_turmas(calendario: dict):
    """Continua existindo para o seletor de turma (só referência/organização)."""
    turmas = []
    mapa = {}
    for nome in calendario.get("turmas_simples", {}):
        turmas.append(nome)
        mapa[nome] = ("simples", nome)
    for nome in calendario.get("turmas_contraturno", {}):
        disp = f"{nome} (contraturno)"
        turmas.append(disp)
        mapa[disp] = ("contraturno", nome)
    return turmas, mapa


# ----------------------------------------------------------------------------
# Normalização de nomes
# ----------------------------------------------------------------------------
def normalizar_nome(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    for ch in [":", "-", "_", ".", ",", ";", "/", "\\", "(", ")", "[", "]"]:
        s = s.replace(ch, " ")
    s = s.replace(" anos", " ano")
    return " ".join(s.split())


def normalizar_nome_aluno(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.upper()
    s = re.sub(r"[^A-Z ]", " ", s)
    return " ".join(s.split())


def encontrar_turma_correspondente(valor, turmas_validas):
    if valor is None:
        return None
    alvo = normalizar_nome(valor)
    if not alvo:
        return None
    for t in turmas_validas:
        if normalizar_nome(t) == alvo:
            return t
    for t in turmas_validas:
        nt = normalizar_nome(t)
        if alvo in nt or nt in alvo:
            return t
    return None


# ----------------------------------------------------------------------------
# Cálculo (nova lógica: por dias)
# ----------------------------------------------------------------------------
def calcular_frequencia(dias_letivos: int, faltas: int):
    """% Frequência = 1 − (faltas ÷ dias letivos). Retorna None se dias_letivos=0."""
    if not dias_letivos:
        return None
    return max(0.0, 1 - (faltas / dias_letivos))


# ----------------------------------------------------------------------------
# Preenchimento do ofício .docx (inalterado)
# ----------------------------------------------------------------------------
def _iter_blocos_documento(doc):
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield {"tipo": "paragrafo", "conteudo": Paragraph(child, doc)}
        elif child.tag.endswith("}tbl"):
            yield {"tipo": "tabela", "conteudo": Table(child, doc)}


def _extrair_nome_do_paragrafo(texto: str):
    m = re.match(r"^\s*Nome\s*:\s*(.+?)\s*$", texto, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return None


def _formatar_percentual(valor: float) -> str:
    return f"{valor:.2f}".replace(".", ",") + "%"


def _preencher_tabela_aluno(tabela, valores_por_mes: dict, sobrescrever: bool = True):
    preenchidos = []
    for row in tabela.rows:
        if len(row.cells) < 2:
            continue

        mes_encontrado = None
        idx_mes = None
        for idx_cel, cell in enumerate(row.cells):
            txt = cell.text.strip().lower()
            for mes in valores_por_mes.keys():
                if txt == mes.lower():
                    mes_encontrado = mes
                    idx_mes = idx_cel
                    break
            if mes_encontrado:
                break

        if mes_encontrado is None:
            continue

        idx_alvo = idx_mes + 1
        if idx_alvo >= len(row.cells):
            continue

        celula = row.cells[idx_alvo]
        if celula.text.strip() and not sobrescrever:
            continue

        valor = valores_por_mes[mes_encontrado]
        if valor is None:
            continue

        celula.text = _formatar_percentual(valor)
        preenchidos.append((mes_encontrado, valor))

    return preenchidos


def preencher_oficio_docx(arquivo_docx, df_resultado: pd.DataFrame, sobrescrever: bool = True):
    from docx import Document

    doc = Document(arquivo_docx)

    mapa_alunos = {}
    for _, row in df_resultado.iterrows():
        nome_norm = normalizar_nome_aluno(row["Aluno"])
        if not nome_norm:
            continue
        if row["% Frequência"] is None or pd.isna(row["% Frequência"]):
            continue
        mapa_alunos.setdefault(nome_norm, {})[row["Mês"]] = float(row["% Frequência"])

    relatorio = {"preenchidos": [], "nao_encontrados": [], "sem_valor": []}

    ultimo_nome = None
    ultimo_nome_norm = None

    for bloco in _iter_blocos_documento(doc):
        if bloco["tipo"] == "paragrafo":
            nome = _extrair_nome_do_paragrafo(bloco["conteudo"].text)
            if nome:
                ultimo_nome = nome
                ultimo_nome_norm = normalizar_nome_aluno(nome)
        elif bloco["tipo"] == "tabela":
            if ultimo_nome_norm is None:
                continue

            valores_por_mes = mapa_alunos.get(ultimo_nome_norm)
            if not valores_por_mes:
                for chave, vals in mapa_alunos.items():
                    if ultimo_nome_norm in chave or chave in ultimo_nome_norm:
                        valores_por_mes = vals
                        break

            if not valores_por_mes:
                if ultimo_nome not in relatorio["nao_encontrados"]:
                    relatorio["nao_encontrados"].append(ultimo_nome)
                continue

            preenchidos = _preencher_tabela_aluno(
                bloco["conteudo"], valores_por_mes, sobrescrever=sobrescrever
            )
            if preenchidos:
                relatorio["preenchidos"].append((ultimo_nome, preenchidos))
            else:
                if ultimo_nome not in relatorio["sem_valor"]:
                    relatorio["sem_valor"].append(ultimo_nome)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue(), relatorio


# ----------------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------------
garantir_calendario_na_sessao()
calendario = st.session_state["calendario"]
turmas, mapa_turmas = lista_turmas(calendario)
meses = calendario["meses"]

st.title("📋 Calculadora de Frequência Escolar — Bolsa Família")
st.caption(
    "Calcula o percentual de frequência a partir dos **dias letivos** de cada mês "
    "e do **número de faltas (em dias)**. Fórmula: % Frequência = 1 − (faltas ÷ dias letivos)."
)

aba_individual, aba_lote, aba_oficio, aba_calendario = st.tabs(
    [
        "👤 Cálculo individual",
        "👥 Cálculo em lote (vários alunos)",
        "📄 Preencher Ofício (.docx)",
        "🗓️ Calendário / Configurações",
    ]
)

# ----------------------------------------------------------------------------
# Aba 1: cálculo individual
# ----------------------------------------------------------------------------
with aba_individual:
    if not turmas:
        st.warning(
            "Nenhuma turma cadastrada ainda. Vá até a aba **Calendário / Configurações** "
            "para adicionar turmas (a turma aqui é só referência/organização, não "
            "entra no cálculo)."
        )
    else:
        col1, col2, col3 = st.columns(3)
        with col1:
            turma_sel = st.selectbox("Turma / Ano escolar", turmas, key="ind_turma")
        with col2:
            mes_sel = st.selectbox("Mês", meses, key="ind_mes")
        with col3:
            faltas_sel = st.number_input(
                "Nº de faltas (em dias) no mês",
                min_value=0, max_value=31, value=0, step=1,
                key="ind_faltas",
            )

        dias_calendario = int(calendario["dias_letivos_por_mes"].get(mes_sel, 0))

        usar_manual = st.checkbox(
            "✏️ Inserir os dias letivos deste mês manualmente (só para esta consulta)",
            value=False,
            key="ind_usar_manual",
            help=(
                "Marque se, só para esta consulta, o mês tem um número de dias letivos "
                "diferente do que está salvo no calendário. O valor do calendário não "
                "é alterado."
            ),
        )

        if usar_manual:
            dias_letivos = st.number_input(
                "Dias letivos neste mês",
                min_value=0, max_value=31, value=dias_calendario, step=1,
                key="ind_dias_manual",
            )
        else:
            dias_letivos = dias_calendario
            st.info(
                f"📅 Usando o calendário: **{dias_letivos} dias letivos** em **{mes_sel}**. "
                "Para alterar de forma permanente, edite na aba **Calendário / Configurações**."
            )

        freq = calcular_frequencia(int(dias_letivos), int(faltas_sel))

        st.divider()

        c1, c2 = st.columns(2)
        c1.metric("Dias letivos no mês", int(dias_letivos))
        c2.metric("Faltas (dias)", int(faltas_sel))

        st.divider()

        if freq is not None:
            pct = freq * 100
            st.subheader(f"Percentual de frequência: {pct:.2f}%")
            if freq < LIMITE_FREQUENCIA:
                st.error(
                    f"⚠️ Frequência abaixo de {LIMITE_FREQUENCIA*100:.0f}% "
                    "(mínimo exigido pelo Bolsa Família)."
                )
            else:
                st.success("✅ Frequência dentro do mínimo exigido pelo Bolsa Família.")
        else:
            st.warning("Não foi possível calcular — informe dias letivos maiores que zero.")

# ----------------------------------------------------------------------------
# Aba 2: cálculo em lote
# ----------------------------------------------------------------------------
with aba_lote:
    st.write(
        "Preencha a tabela abaixo com **Aluno**, **Turma**, **Mês** e **Faltas**, "
        "ou importe um arquivo CSV/Excel com essas colunas."
    )
    st.caption(
        "💡 A coluna **Turma** continua existindo só para organização/referência — "
        "não entra no cálculo. O cálculo usa **Mês + Faltas (em dias)** e os dias "
        "letivos do calendário. A coluna **Dias letivos (opcional)** permite "
        "sobrescrever o valor do calendário só naquela linha. "
        "Para apagar uma linha, clique nela e aperte **Delete/Backspace**."
    )

    colunas_lote = ["Aluno", "Turma", "Mês", "Faltas", "Dias letivos"]

    def tabela_vazia():
        return pd.DataFrame(
            {
                "Aluno": pd.Series(dtype="object"),
                "Turma": pd.Series(dtype="object"),
                "Mês": pd.Series(dtype="object"),
                "Faltas": pd.Series(dtype="int64"),
                "Dias letivos": pd.Series(dtype="Int64"),
            }
        )

    def tabela_modelo():
        return pd.DataFrame(
            {
                "Aluno": ["Exemplo: João da Silva"],
                "Turma": [turmas[0] if turmas else ""],
                "Mês": [meses[0]],
                "Faltas": [0],
                "Dias letivos": [pd.NA],
            }
        )

    if "tabela_lote" not in st.session_state:
        st.session_state["tabela_lote"] = tabela_modelo()

    col_upload, col_limpar = st.columns([3, 1])
    with col_upload:
        arquivo = st.file_uploader(
            "Importar planilha (opcional) — colunas: Aluno, Turma, Mês, Faltas [, Dias letivos]",
            type=["csv", "xlsx"],
        )
    with col_limpar:
        st.write("")
        st.write("")
        if st.button("🗑️ Limpar tabela", use_container_width=True):
            st.session_state["tabela_lote"] = tabela_vazia()
            st.session_state.pop("df_resultado", None)
            st.rerun()

    if arquivo is not None:
        try:
            if arquivo.name.endswith(".csv"):
                df_importado = pd.read_csv(arquivo)
            else:
                df_importado = pd.read_excel(arquivo)

            colunas_minimas = {"Aluno", "Turma", "Mês", "Faltas"}
            if not colunas_minimas.issubset(set(df_importado.columns)):
                st.error(
                    f"O arquivo precisa conter no mínimo as colunas: "
                    f"{', '.join(sorted(colunas_minimas))}. "
                    "A coluna 'Dias letivos' é opcional."
                )
            else:
                if "Dias letivos" not in df_importado.columns:
                    df_importado["Dias letivos"] = pd.NA
                df_importado = df_importado[
                    ["Aluno", "Turma", "Mês", "Faltas", "Dias letivos"]
                ].copy()

                # Normaliza o nome da turma (tolerante a acentos/maiúsculas)
                nao_reconhecidas = []
                turmas_corrigidas = []
                for valor in df_importado["Turma"]:
                    corr = encontrar_turma_correspondente(valor, turmas)
                    turmas_corrigidas.append(corr if corr is not None else valor)
                    if corr is None and not pd.isna(valor):
                        nao_reconhecidas.append(str(valor))
                df_importado["Turma"] = turmas_corrigidas

                st.session_state["tabela_lote"] = df_importado

                if nao_reconhecidas:
                    st.warning(
                        "⚠️ Algumas turmas da planilha não foram reconhecidas "
                        "automaticamente (não afeta o cálculo, é só a coluna de "
                        "referência). Você pode corrigir na tabela se quiser:\n\n"
                        + "\n".join(f"- `{v}`" for v in sorted(set(nao_reconhecidas)))
                    )
                else:
                    st.success("Arquivo importado com sucesso!")
        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")

    tabela_editada = st.data_editor(
        st.session_state["tabela_lote"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Turma": st.column_config.SelectboxColumn(
                "Turma (referência)", options=turmas, required=False,
                help="Não entra no cálculo — serve só para você se organizar."
            ),
            "Mês": st.column_config.SelectboxColumn("Mês", options=meses, required=True),
            "Faltas": st.column_config.NumberColumn(
                "Faltas (em dias)", min_value=0, max_value=31, step=1, required=True,
                help="Faltas em dias (não por disciplina).",
            ),
            "Dias letivos": st.column_config.NumberColumn(
                "Dias letivos (opcional)",
                min_value=0, max_value=31, step=1,
                help="Deixe em branco para usar o valor do calendário. "
                     "Preencha só se este aluno/mês tiver um valor diferente.",
            ),
        },
        key="editor_lote",
    )
    st.session_state["tabela_lote"] = tabela_editada

    col_b1, col_b2 = st.columns([1, 3])
    with col_b1:
        if st.button("➖ Apagar última linha", use_container_width=True):
            if len(st.session_state["tabela_lote"]) > 0:
                st.session_state["tabela_lote"] = (
                    st.session_state["tabela_lote"].iloc[:-1].reset_index(drop=True)
                )
                st.session_state.pop("df_resultado", None)
                st.rerun()
    with col_b2:
        st.caption(
            "Use este botão se o ícone de lixeira do editor não estiver aparecendo "
            "na sua versão do Streamlit."
        )

    if st.button("Calcular frequência de todos os alunos", type="primary"):
        linhas = []
        problemas = []
        for _, row in tabela_editada.iterrows():
            if pd.isna(row.get("Mês")):
                continue
            mes = row["Mês"]

            # Dias letivos: manual (linha) > calendário
            dias_manual = row.get("Dias letivos")
            if dias_manual is not None and not pd.isna(dias_manual):
                dias_letivos = int(dias_manual)
                origem_dias = "manual (linha)"
            else:
                dias_letivos = int(calendario["dias_letivos_por_mes"].get(mes, 0))
                origem_dias = "calendário"

            faltas = int(row["Faltas"]) if not pd.isna(row["Faltas"]) else 0

            if dias_letivos == 0:
                problemas.append(
                    f"{row['Aluno']} — mês `{mes}` sem dias letivos definidos"
                )
                freq = None
            else:
                freq = calcular_frequencia(dias_letivos, faltas)

            linhas.append(
                {
                    "Aluno": row["Aluno"],
                    "Turma": row.get("Turma"),
                    "Mês": mes,
                    "Dias letivos no mês": dias_letivos,
                    "Origem dos dias": origem_dias,
                    "Faltas (dias)": faltas,
                    "% Frequência": round(freq * 100, 2) if freq is not None else None,
                    "Abaixo de 75%": (
                        "⚠️ Sim"
                        if (freq is not None and freq < LIMITE_FREQUENCIA)
                        else ("❓ Verificar mês" if freq is None else "Não")
                    ),
                }
            )

        if problemas:
            st.error(
                "❌ Algumas linhas não puderam ser calculadas:\n\n"
                + "\n".join(f"- {p}" for p in problemas)
            )

        if linhas:
            df_resultado = pd.DataFrame(linhas)
            st.session_state["df_resultado"] = df_resultado

    if "df_resultado" in st.session_state:
        st.divider()
        st.subheader("Resultado")

        def destacar_abaixo(row):
            if row["Abaixo de 75%"] == "⚠️ Sim":
                return ["background-color: #ffe0e0"] * len(row)
            if row["Abaixo de 75%"] == "❓ Verificar mês":
                return ["background-color: #fff4cc"] * len(row)
            return [""] * len(row)

        st.dataframe(
            st.session_state["df_resultado"].style.apply(destacar_abaixo, axis=1),
            use_container_width=True,
        )

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            st.session_state["df_resultado"].to_excel(
                writer, index=False, sheet_name="Frequência"
            )
        st.download_button(
            "⬇️ Baixar resultado em Excel",
            data=buffer.getvalue(),
            file_name="frequencia_bolsa_familia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ----------------------------------------------------------------------------
# Aba 3: preencher ofício .docx (inalterada)
# ----------------------------------------------------------------------------
with aba_oficio:
    st.write(
        "Faça upload do **ofício .docx** recebido (o mesmo modelo do Bolsa Família) "
        "e o app preenche automaticamente as colunas **Frequência** de cada mês "
        "com os percentuais calculados na aba **Cálculo em lote**."
    )

    if "df_resultado" not in st.session_state:
        st.warning(
            "⚠️ Primeiro vá na aba **👥 Cálculo em lote** e clique em "
            "**\"Calcular frequência de todos os alunos\"**. "
            "O preenchimento do ofício usa os resultados daquela aba."
        )
    else:
        st.success(
            f"✅ {len(st.session_state['df_resultado'])} registros de frequência "
            "disponíveis para preencher o ofício."
        )

        arquivo_docx = st.file_uploader(
            "Upload do ofício (.docx)", type=["docx"], key="upload_oficio"
        )

        sobrescrever = st.checkbox(
            "Sobrescrever valores já existentes nas células",
            value=False,
            help="Se desmarcado, o app só preenche células vazias. "
                 "Se marcado, substitui o que já estiver lá.",
        )

        if arquivo_docx is not None:
            if st.button("🚀 Preencher ofício e gerar arquivo", type="primary"):
                try:
                    with st.spinner("Preenchendo o ofício..."):
                        bytes_docx, relatorio = preencher_oficio_docx(
                            arquivo_docx,
                            st.session_state["df_resultado"],
                            sobrescrever=sobrescrever,
                        )
                    st.session_state["docx_preenchido"] = bytes_docx

                    st.success("Ofício processado!")

                    c1, c2, c3 = st.columns(3)
                    c1.metric("Alunos preenchidos", len(relatorio["preenchidos"]))
                    c2.metric("Não encontrados", len(relatorio["nao_encontrados"]))
                    c3.metric("Sem valor calculado", len(relatorio["sem_valor"]))

                    if relatorio["preenchidos"]:
                        with st.expander(
                            f"✅ {len(relatorio['preenchidos'])} aluno(s) preenchidos",
                            expanded=False,
                        ):
                            for nome, itens in relatorio["preenchidos"]:
                                detalhes = ", ".join(
                                    f"{mes}={_formatar_percentual(val)}"
                                    for mes, val in itens
                                )
                                st.write(f"- **{nome}** → {detalhes}")

                    if relatorio["nao_encontrados"]:
                        with st.expander(
                            f"⚠️ {len(relatorio['nao_encontrados'])} aluno(s) do ofício "
                            "NÃO encontrados na planilha",
                            expanded=True,
                        ):
                            st.caption(
                                "Esses nomes apareceram no .docx mas não têm cálculo na "
                                "aba de lote. Confira se digitou o nome igual à planilha."
                            )
                            for n in relatorio["nao_encontrados"]:
                                st.write(f"- {n}")

                    if relatorio["sem_valor"]:
                        with st.expander(
                            f"⚠️ {len(relatorio['sem_valor'])} aluno(s) encontrados "
                            "mas sem % calculado",
                            expanded=True,
                        ):
                            st.caption(
                                "O nome foi achado, mas não havia valor de frequência "
                                "para os meses da tabela (ou o mês não tem dias letivos)."
                            )
                            for n in relatorio["sem_valor"]:
                                st.write(f"- {n}")

                except Exception as e:
                    st.error(f"Erro ao processar o .docx: {e}")
                    st.exception(e)

        if "docx_preenchido" in st.session_state:
            st.download_button(
                "⬇️ Baixar ofício preenchido (.docx)",
                data=st.session_state["docx_preenchido"],
                file_name="oficio_preenchido.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

# ----------------------------------------------------------------------------
# Aba 4: calendário / configurações (mantendo as 3 subtabs)
# ----------------------------------------------------------------------------
with aba_calendario:
    st.write(
        "Aqui você pode editar **tudo** o que o cálculo usa: os dias letivos de cada mês, "
        "as turmas regulares e as turmas com contraturno — direto pela tela, sem precisar "
        "mexer em nenhum arquivo. As alterações ficam valendo nesta sessão; use o botão "
        "**\"💾 Salvar no arquivo\"** para gravar de vez, ou **\"⬇️ Baixar cópia (.json)\"** "
        "para guardar uma cópia de segurança no seu computador."
    )
    st.info(
        "ℹ️ **Importante:** com a nova regra por **dias letivos**, apenas a subtab "
        "**📅 Dias letivos por mês** entra no cálculo. As subtabs **🏫 Turmas regulares** "
        "e **🕑 Turmas com contraturno** continuam aqui para consulta/registro, mas "
        "**não são mais usadas** no cálculo de frequência."
    )

    with st.expander("📤 Restaurar calendário a partir de um arquivo .json", expanded=False):
        st.caption(
            "Se você baixou uma cópia antes e o app perdeu as alterações, suba aqui o arquivo "
            "para restaurar tudo de uma vez. Depois clique em **\"💾 Salvar no arquivo\"** para gravar."
        )
        json_upload = st.file_uploader(
            "Suba o arquivo calendario_2026.json", type=["json"], key="upload_calendario"
        )
        if json_upload is not None:
            try:
                novo_calendario = json.load(json_upload)
                if "dias_letivos_por_mes" not in novo_calendario:
                    st.error(
                        "O arquivo não parece ser um calendário válido "
                        "(falta a chave 'dias_letivos_por_mes')."
                    )
                else:
                    st.session_state["calendario"] = novo_calendario
                    st.success(
                        "Calendário carregado! Clique em 'Salvar no arquivo' abaixo "
                        "para gravar no servidor."
                    )
                    st.rerun()
            except Exception as e:
                st.error(f"Erro ao ler o JSON: {e}")

    sub_dias, sub_simples, sub_contra = st.tabs(
        ["📅 Dias letivos por mês", "🏫 Turmas regulares", "🕑 Turmas com contraturno"]
    )

    # ----- Subtabs 1: dias letivos por mês (é o que realmente importa) -----
    with sub_dias:
        st.caption(
            "Edite o número de dias letivos de cada mês. **Isso é o que entra no cálculo.** "
            "Cada mês pode ter um valor diferente."
        )
        df_dias = pd.DataFrame(
            {
                "Mês": meses,
                "Dias letivos": [
                    calendario["dias_letivos_por_mes"].get(m, 0) for m in meses
                ],
            }
        )
        df_dias_editado = st.data_editor(
            df_dias,
            use_container_width=True,
            hide_index=True,
            disabled=["Mês"],
            column_config={
                "Dias letivos": st.column_config.NumberColumn(
                    "Dias letivos", min_value=0, max_value=31, step=1
                )
            },
            key="editor_dias_letivos",
        )
        if st.button("Aplicar dias letivos", key="btn_aplicar_dias", type="primary"):
            for _, row in df_dias_editado.iterrows():
                calendario["dias_letivos_por_mes"][row["Mês"]] = int(row["Dias letivos"])
            st.success(
                "Dias letivos atualizados! (lembre de salvar no arquivo, se quiser manter)"
            )
            st.rerun()

        total_ano = sum(
            calendario["dias_letivos_por_mes"].get(m, 0) for m in meses
        )
        st.metric("Total de dias letivos no ano", total_ano)

    # ----- Subtab 2: turmas regulares (agora só referência) -----
    with sub_simples:
        st.caption(
            "⚠️ Esta subtab **não entra mais no cálculo** (que agora é por dias). "
            "Mantenha aqui as turmas só para registro/consulta. Você pode adicionar, "
            "renomear ou remover turmas normalmente."
        )
        df_simples = pd.DataFrame(
            [
                {"Turma": k, "Aulas por dia": v.get("aulas_por_dia", 0)}
                for k, v in calendario.get("turmas_simples", {}).items()
            ]
        )
        df_simples_editado = st.data_editor(
            df_simples,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Aulas por dia": st.column_config.NumberColumn(
                    "Aulas por dia", min_value=0, max_value=20, step=1
                )
            },
            key="editor_turmas_simples",
        )
        if st.button("Aplicar turmas regulares", key="btn_aplicar_simples"):
            novas = {}
            for _, row in df_simples_editado.iterrows():
                nome = str(row.get("Turma") or "").strip()
                if not nome or pd.isna(row.get("Aulas por dia")):
                    continue
                novas[nome] = {"aulas_por_dia": int(row["Aulas por dia"])}
            calendario["turmas_simples"] = novas
            st.success("Turmas regulares atualizadas!")
            st.rerun()

    # ----- Subtab 3: turmas com contraturno (agora só referência) -----
    with sub_contra:
        st.caption(
            "⚠️ Esta subtab **não entra mais no cálculo** (que agora é por dias). "
            "Mantenha aqui as turmas com contraturno só para registro/consulta."
        )

        turmas_contra = calendario.setdefault("turmas_contraturno", {})

        with st.expander("➕ Adicionar nova turma com contraturno"):
            novo_nome = st.text_input(
                "Nome/código da turma (ex.: 21301, 3ºC, etc.)", key="novo_nome_contra"
            )
            if st.button("Adicionar turma", key="btn_add_contra"):
                nome_limpo = novo_nome.strip()
                if not nome_limpo:
                    st.warning("Digite um nome para a turma.")
                elif nome_limpo in turmas_contra:
                    st.warning("Já existe uma turma com esse nome.")
                else:
                    turmas_contra[nome_limpo] = {
                        "aulas_manha_por_dia": 0,
                        "aulas_por_dia_semana": {d: 0 for d in DIAS_SEMANA},
                        "dias_letivos_por_dia_semana_por_mes": {
                            m: {d: 0 for d in DIAS_SEMANA} for m in meses
                        },
                    }
                    st.success(
                        f"Turma '{nome_limpo}' adicionada. Configure os detalhes dela abaixo."
                    )
                    st.rerun()

        if not turmas_contra:
            st.info("Nenhuma turma com contraturno cadastrada ainda.")

        for nome_turma in list(turmas_contra.keys()):
            info = turmas_contra[nome_turma]
            with st.expander(f"🕑 {nome_turma}", expanded=False):
                col_nome, col_remover = st.columns([3, 1])
                with col_nome:
                    novo_nome_turma = st.text_input(
                        "Nome/código da turma (pode editar para renomear)",
                        value=nome_turma,
                        key=f"nome_{nome_turma}",
                    )
                with col_remover:
                    st.write("")
                    st.write("")
                    if st.button("🗑️ Remover turma", key=f"remover_{nome_turma}"):
                        del turmas_contra[nome_turma]
                        st.success(f"Turma '{nome_turma}' removida.")
                        st.rerun()

                aulas_manha = st.number_input(
                    "Aulas pela manhã por dia letivo",
                    min_value=0, max_value=20, step=1,
                    value=int(info.get("aulas_manha_por_dia", 0)),
                    key=f"manha_{nome_turma}",
                )

                st.markdown("**Aulas no contraturno, por dia da semana**")
                st.caption("Deixe 0 nos dias em que a turma não tem contraturno.")
                aulas_semana_atual = info.get(
                    "aulas_por_dia_semana", {d: 0 for d in DIAS_SEMANA}
                )
                cols_semana = st.columns(len(DIAS_SEMANA))
                novos_valores_semana = {}
                for col, dia in zip(cols_semana, DIAS_SEMANA):
                    with col:
                        novos_valores_semana[dia] = st.number_input(
                            DIAS_SEMANA_LABEL[dia],
                            min_value=0, max_value=20, step=1,
                            value=int(aulas_semana_atual.get(dia, 0)),
                            key=f"semana_{nome_turma}_{dia}",
                        )

                st.markdown(
                    "**Dias letivos de contraturno por mês (quantos de cada dia da semana)**"
                )
                dias_mes_atual = info.get("dias_letivos_por_dia_semana_por_mes", {})
                df_mensal = pd.DataFrame(
                    [
                        {
                            "Mês": m,
                            **{
                                DIAS_SEMANA_LABEL[d]: dias_mes_atual.get(m, {}).get(d, 0)
                                for d in DIAS_SEMANA
                            },
                        }
                        for m in meses
                    ]
                )
                col_config_mensal = {
                    DIAS_SEMANA_LABEL[d]: st.column_config.NumberColumn(
                        DIAS_SEMANA_LABEL[d], min_value=0, max_value=6, step=1
                    )
                    for d in DIAS_SEMANA
                }
                df_mensal_editado = st.data_editor(
                    df_mensal,
                    use_container_width=True,
                    hide_index=True,
                    disabled=["Mês"],
                    column_config=col_config_mensal,
                    key=f"editor_mensal_{nome_turma}",
                )

                if st.button(
                    "💾 Aplicar alterações desta turma", key=f"aplicar_{nome_turma}"
                ):
                    nome_final = novo_nome_turma.strip()
                    if not nome_final:
                        st.error("O nome da turma não pode ficar vazio.")
                    else:
                        if nome_final != nome_turma:
                            if nome_final in turmas_contra:
                                st.error(
                                    f"Já existe uma turma chamada '{nome_final}'. "
                                    "Escolha outro nome."
                                )
                                st.stop()
                            turmas_contra[nome_final] = turmas_contra.pop(nome_turma)
                            info = turmas_contra[nome_final]

                        info["aulas_manha_por_dia"] = int(aulas_manha)
                        info["aulas_por_dia_semana"] = {
                            d: int(novos_valores_semana[d]) for d in DIAS_SEMANA
                        }
                        novo_mensal = {}
                        for _, row in df_mensal_editado.iterrows():
                            novo_mensal[row["Mês"]] = {
                                d: int(row[DIAS_SEMANA_LABEL[d]]) for d in DIAS_SEMANA
                            }
                        info["dias_letivos_por_dia_semana_por_mes"] = novo_mensal
                        st.success(f"Turma '{nome_final}' atualizada!")
                        st.rerun()

    st.divider()
    col_salvar, col_baixar = st.columns(2)
    with col_salvar:
        if st.button(
            "💾 Salvar no arquivo (calendario_2026.json)",
            type="primary",
            use_container_width=True,
        ):
            try:
                salvar_calendario_no_disco()
                st.success("Configurações salvas no arquivo calendario_2026.json!")
            except Exception as e:
                st.error(
                    f"Não foi possível salvar no arquivo do servidor ({e}). "
                    "Use o botão 'Baixar cópia' ao lado para guardar suas alterações."
                )
    with col_baixar:
        st.download_button(
            "⬇️ Baixar cópia (.json)",
            data=json.dumps(calendario, ensure_ascii=False, indent=2),
            file_name="calendario_2026.json",
            mime="application/json",
            use_container_width=True,
        )
    st.caption(
        "⚠️ Em alguns tipos de hospedagem (como o Streamlit Community Cloud), o arquivo salvo no "
        "servidor pode ser perdido quando o app reinicia ou é atualizado via GitHub. Por segurança, "
        "baixe uma cópia do JSON de vez em quando e, se precisar, use o uploader acima para restaurar."
    )

st.divider()
st.caption(
    "⚠️ O percentual mínimo de frequência exigido pelo Bolsa Família (condicionalidade de educação) "
    "é geralmente de 75%. Confirme sempre as regras vigentes com a coordenação/Ministério, pois "
    "podem mudar."
)