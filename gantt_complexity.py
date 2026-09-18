from __future__ import annotations

import io
import math
import re
import subprocess
import sys
import tempfile
from collections import defaultdict
from pathlib import Path

import networkx as nx
import pandas as pd
import plotly.graph_objects as go


def _avvia_streamlit_se_necessario() -> None:
    """Consente anche l'avvio con: python gantt_complexity_streamlit.py"""
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        dentro_streamlit = get_script_run_ctx() is not None
    except Exception:
        dentro_streamlit = False

    if not dentro_streamlit:
        subprocess.run(
            [sys.executable, "-m", "streamlit", "run", str(Path(__file__).resolve())],
            check=False,
        )
        raise SystemExit


_avvia_streamlit_se_necessario()

import streamlit as st


##

import importlib.metadata
import importlib.util
import sys

with st.sidebar.expander("Aspose diagnostics"):
    try:
        distribution = importlib.metadata.distribution("aspose-tasks")

        st.write("Distribution found:", distribution.metadata["Name"])
        st.write("Version:", distribution.version)

        installed_files = [
            str(file)
            for file in (distribution.files or [])
            if "aspose" in str(file).lower()
        ]

        st.write("Installed files:")
        st.code("\n".join(installed_files[:100]))

    except Exception as exc:
        st.error(
            f"Distribution check failed: "
            f"{type(exc).__name__}: {exc}"
        )

    st.write(
        "aspose namespace:",
        importlib.util.find_spec("aspose")
    )

    st.write(
        "aspose.tasks module:",
        importlib.util.find_spec("aspose.tasks")
        if importlib.util.find_spec("aspose") is not None
        else None
    )

    st.write("Python executable:", sys.executable)

import platform
import sys

st.write("Python version:", sys.version)
st.write("Python executable:", sys.executable)
st.write("Platform:", platform.platform())

##

COL_MACRO = "PrismaMacroActivity"
COL_ID = "ID"
COL_PREDECESSORI = "Predecessori"
COL_NOME = "Nome"
COL_DURATA = "Durata"
COL_INIZIO = "Inizio"
COL_FINE = "Fine"
COL_SUMMARY = "IsSummary"

REQUIRED_COLUMNS = {
    COL_MACRO,
    COL_ID,
    COL_PREDECESSORI,
    COL_NOME,
    COL_DURATA,
    COL_INIZIO,
    COL_FINE,
}

RELATION_COLORS = {
    "FI": "#64748b",
    "FS": "#64748b",
    "II": "#2563eb",
    "FF": "#7c3aed",
    "IF": "#ea580c",
}


def normalizza_id(value) -> str | None:
    if pd.isna(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return text or None


def estrai_predecessori(value) -> tuple[list[dict], list[str]]:
    if pd.isna(value) or not str(value).strip():
        return [], []

    risultati = []
    non_riconosciuti = []
    pattern = re.compile(
        r"^(?P<id>\d+)\s*(?P<rel>FI|FF|II|IF)?\s*"
        r"(?P<lag>[+-]\s*\d+(?:[.,]\d+)?)?\s*g?\s*$",
        flags=re.IGNORECASE,
    )

    for token in str(value).split(";"):
        token = token.strip()
        if not token:
            continue
        match = pattern.match(token)
        if not match:
            non_riconosciuti.append(token)
            continue
        risultati.append(
            {
                "id": match.group("id"),
                "relazione": (match.group("rel") or "FI").upper(),
                "lag": float(
                    (match.group("lag") or "0")
                    .replace(" ", "")
                    .replace(",", ".")
                ),
                "testo_originale": token,
            }
        )
    return risultati, non_riconosciuti


def _data_italiana(value):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (pd.Timestamp,)):
        return value
    testo = str(value).strip()
    testo = re.sub(r"^[A-Za-zÀ-ÿ]{3}\s+", "", testo)
    return pd.to_datetime(testo, errors="coerce", dayfirst=True)


@st.cache_data(show_spinner=False)
def leggi_fogli(contenuto: bytes) -> list[str]:
    return pd.ExcelFile(io.BytesIO(contenuto), engine="openpyxl").sheet_names


@st.cache_data(show_spinner=False)
def carica_schedule(contenuto: bytes, foglio: str) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(contenuto), sheet_name=foglio, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    mancanti = REQUIRED_COLUMNS.difference(df.columns)
    if mancanti:
        raise ValueError("Missing columns: " + ", ".join(sorted(mancanti)))

    df = df.copy()
    df[COL_ID] = df[COL_ID].apply(normalizza_id)
    df = df[df[COL_ID].notna()].copy()
    df[COL_NOME] = df[COL_NOME].fillna("").astype(str).str.strip()
    df[COL_MACRO] = (
        df[COL_MACRO].fillna("Unclassified").astype(str).str.strip()
    )
    df.loc[df[COL_MACRO] == "", COL_MACRO] = "Unclassified"
    df[COL_INIZIO] = df[COL_INIZIO].apply(_data_italiana)
    df[COL_FINE] = df[COL_FINE].apply(_data_italiana)
    colonna_summary = next(
        (col for col in [COL_SUMMARY, "Summary", "Riepilogo", "Tasks di riepilogo", "Attivita di riepilogo"] if col in df.columns),
        None,
    )
    if colonna_summary is None:
        df[COL_SUMMARY] = False
    else:
        valori_veri = {"true", "vero", "yes", "si", "sì", "1", "x"}
        df[COL_SUMMARY] = df[colonna_summary].fillna(False).apply(
            lambda valore: str(valore).strip().lower() in valori_veri
        )
    # Le attività di riepilogo vengono eliminate direttamente al caricamento.
    numero_summary_escluse = int(df[COL_SUMMARY].sum())
    df = df.loc[~df[COL_SUMMARY]].copy()
    df.drop(columns=[COL_SUMMARY], inplace=True, errors="ignore")
    df.attrs["summary_escluse"] = numero_summary_escluse
    return df


def testo_enum(value) -> str:
    return str(value).split(".")[-1].upper().replace(" ", "_")


def codice_relazione(link_type) -> str:
    """Converte TaskLinkType Aspose, che in Python viene esposto come intero."""
    valore = str(link_type).strip()
    mapping_numerico = {
        "0": "FF",  # End nodessh-to-End nodessh
        "1": "FI",  # End nodessh-to-Start
        "2": "IF",  # Start-to-End nodessh
        "3": "II",  # Start-to-Start
    }
    if valore in mapping_numerico:
        return mapping_numerico[valore]

    nome = testo_enum(link_type)
    return {
        "FINISH_TO_START": "FI",
        "START_TO_START": "II",
        "FINISH_TO_FINISH": "FF",
        "START_TO_FINISH": "IF",
    }.get(nome, nome)


def numero_lag_giorni(link) -> float:
    """
    Converte il lag Aspose in giorni lavorativi Project.
    Nel file diagnosticato lag_format=4 (giorni) e link_lag è espresso
    in decimi di minuto: 4.800 unità corrispondono a 1 giorno da 8 ore.
    """
    try:
        lag_raw = float(link.link_lag)
        lag_format = str(link.lag_format).strip()
        if lag_format == "4":
            return lag_raw / 4800.0
    except Exception:
        pass

    # Fallback per formati diversi dai giorni.
    try:
        return float(link.link_lag_time_span.total_seconds()) / 3600.0 / 8.0
    except Exception:
        try:
            return float(link.link_lag) / 4800.0
        except Exception:
            return 0.0


def formatta_lag(giorni: float) -> str:
    if abs(giorni) < 1e-9:
        return ""
    valore = int(giorni) if float(giorni).is_integer() else round(giorni, 2)
    return f"{'+' if giorni > 0 else ''}{valore} g"


def valore_attributo_sicuro(attribute) -> str:
    """Legge solo proprietà compatibili con il tipo concreto dell'attributo."""
    for nome in ("value", "text_value", "numeric_value"):
        try:
            valore = getattr(attribute, nome)
        except (AttributeError, RuntimeError):
            continue
        if valore is not None and str(valore).strip():
            return str(valore).strip()
    return ""


def definizioni_attributi(project) -> dict[str, object]:
    risultato = {}
    try:
        for definition in project.extended_attributes:
            try:
                risultato[str(definition.field_id)] = definition
            except Exception:
                continue
    except Exception:
        pass
    return risultato


def risolvi_lookup(definition, attribute, valore_grezzo: str) -> str:
    """Traduce l'eventuale GUID/valore di lookup nel testo leggibile."""
    if definition is None:
        return valore_grezzo

    try:
        guid_attributo = str(attribute.value_guid or "").strip().lower()
    except Exception:
        guid_attributo = ""

    try:
        for item in definition.value_list:
            try:
                item_guid = str(item.value_guid or "").strip().lower()
            except Exception:
                item_guid = ""

            valori = []
            for nome in ("string_value", "val", "description"):
                try:
                    valore = getattr(item, nome)
                except Exception:
                    continue
                if valore is not None and str(valore).strip():
                    valori.append(str(valore).strip())

            if guid_attributo and item_guid == guid_attributo and valori:
                return valori[0]
            if valore_grezzo and valore_grezzo in valori:
                return valori[0]
    except Exception:
        pass

    return valore_grezzo


def leggi_prisma_macro_activity(project, task) -> str:
    """Cerca Prisma Macro Activity per alias/nome; fallback sul campo Text10."""
    definizioni = definizioni_attributi(project)
    fallback_text10 = ""

    for attribute in task.extended_attributes:
        try:
            field_id = str(attribute.field_id or "")
        except Exception:
            field_id = ""

        definition = definizioni.get(field_id)
        alias = ""
        field_name = ""
        if definition is not None:
            try:
                alias = str(definition.alias or "").strip()
            except Exception:
                pass
            try:
                field_name = str(definition.field_name or "").strip()
            except Exception:
                pass

        chiave = f"{alias} {field_name}".lower().replace(" ", "").replace("_", "")
        e_prisma = "prismamacroactivity" in chiave
        e_text10 = "text10" in field_id.lower() or "text10" in field_name.lower()

        # Non leggere text_value su tutti gli attributi: un attributo Number
        # solleva InvalidOperationException proprio come nell'errore segnalato.
        if not (e_prisma or e_text10):
            continue

        valore = valore_attributo_sicuro(attribute)
        valore = risolvi_lookup(definition, attribute, valore)

        if e_prisma and valore:
            return valore
        if e_text10 and valore and not fallback_text10:
            fallback_text10 = valore

    return fallback_text10


def costruisci_predecessori(project) -> dict[int, list[str]]:
    predecessori = defaultdict(list)
    for link in project.task_links:
        pred = link.pred_task
        succ = link.succ_task
        if pred is None or succ is None:
            continue
        pred_id = int(pred.id)
        succ_id = int(succ.id)
        relazione = codice_relazione(link.link_type)
        lag = formatta_lag(numero_lag_giorni(link))
        testo = str(pred_id) if relazione == "FI" and not lag else f"{pred_id}{relazione}{lag}"
        predecessori[succ_id].append(testo)
    return predecessori


@st.cache_data(show_spinner=False)
def carica_schedule_mpp(contenuto: bytes, nome_file: str) -> pd.DataFrame:
    try:
        import aspose.tasks as tasks
    except Exception as exc:
        raise RuntimeError(
        "Aspose.Tasks is installed but could not be initialized. "
        f"Error type: {type(exc).__name__}. "
        f"Error details: {exc}"
    ) from exc

    percorso_temporaneo = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".mpp", delete=False) as tmp:
            tmp.write(contenuto)
            percorso_temporaneo = Path(tmp.name)

        project = tasks.Project(str(percorso_temporaneo))
        predecessori = costruisci_predecessori(project)
        righe = []
        numero_summary_escluse = 0
        for task in project.root_task.select_all_child_tasks():
            task_id = int(task.id)
            # Aspose include anche la root del progetto (ID 0); l'export Excel
            # originale contiene invece le sole attività ID 1..N.
            if task_id == 0:
                continue
            if bool(task.is_summary):
                numero_summary_escluse += 1
                continue
            righe.append({
                COL_MACRO: leggi_prisma_macro_activity(project, task),
                COL_ID: str(task_id),
                COL_PREDECESSORI: ";".join(predecessori.get(task_id, [])),
                COL_NOME: str(task.name or "").strip(),
                COL_DURATA: str(task.duration or ""),
                COL_INIZIO: pd.to_datetime(task.start, errors="coerce"),
                COL_FINE: pd.to_datetime(task.finish, errors="coerce"),
            })

        df = pd.DataFrame(righe, columns=[
            COL_MACRO, COL_ID, COL_PREDECESSORI, COL_NOME,
            COL_DURATA, COL_INIZIO, COL_FINE,
        ])
        df[COL_MACRO] = df[COL_MACRO].fillna("").astype(str).str.strip()
        df.loc[df[COL_MACRO] == "", COL_MACRO] = "Unclassified"
        df.attrs["summary_escluse"] = numero_summary_escluse
        return df
    except Exception as exc:
        raise RuntimeError(f"Unable to read MPP file '{nome_file}': {exc}") from exc
    finally:
        if percorso_temporaneo is not None:
            try:
                percorso_temporaneo.unlink(missing_ok=True)
            except Exception:
                pass


def crea_excel_schedule_estratto(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="data", index=False)
        ws = writer.sheets["data"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col, width in {
            "A": 32, "B": 10, "C": 45, "D": 60,
            "E": 18, "F": 20, "G": 20,
        }.items():
            ws.column_dimensions[col].width = width
        for cell in ws["F"][1:]:
            cell.number_format = "dd/mm/yyyy hh:mm"
        for cell in ws["G"][1:]:
            cell.number_format = "dd/mm/yyyy hh:mm"
    output.seek(0)
    return output.getvalue()


def costruisci_grafo(
    df: pd.DataFrame,
    ids_esclusi_per_data: set[str] | None = None,
) -> tuple[nx.DiGraph, pd.DataFrame]:
    grafo = nx.DiGraph()
    ids_validi = set(df[COL_ID])
    ids_esclusi_per_data = ids_esclusi_per_data or set()
    anomalie = []

    for _, row in df.iterrows():
        grafo.add_node(
            row[COL_ID],
            nome=row[COL_NOME],
            macro=row[COL_MACRO],
            durata=row[COL_DURATA],
            inizio=row[COL_INIZIO],
            fine=row[COL_FINE],
        )

    for _, row in df.iterrows():
        successore = row[COL_ID]
        predecessori, non_riconosciuti = estrai_predecessori(row[COL_PREDECESSORI])

        for token in non_riconosciuti:
            anomalie.append(
                {
                    "Type": "Unrecognized format",
                    "Successor": successore,
                    "Reference": token,
                }
            )

        for pred in predecessori:
            predecessore = pred["id"]
            if predecessore in ids_esclusi_per_data:
                continue
            if predecessore not in ids_validi:
                anomalie.append(
                    {
                        "Type": "Predecessor ID not found",
                        "Successor": successore,
                        "Reference": pred["testo_originale"],
                    }
                )
                continue
            grafo.add_edge(
                predecessore,
                successore,
                relazione=pred["relazione"],
                lag=pred["lag"],
                testo_originale=pred["testo_originale"],
            )

    return grafo, pd.DataFrame(anomalie)


def conta_cammini_dag(grafo: nx.DiGraph) -> tuple[dict, dict]:
    ordine = list(nx.topological_sort(grafo))
    cammini_da_start = {n: 0 for n in grafo.nodes}
    cammini_verso_fine = {n: 0 for n in grafo.nodes}

    for nodo in ordine:
        predecessori = list(grafo.predecessors(nodo))
        cammini_da_start[nodo] = (
            1
            if not predecessori
            else sum(cammini_da_start[p] for p in predecessori)
        )

    for nodo in reversed(ordine):
        successori = list(grafo.successors(nodo))
        cammini_verso_fine[nodo] = (
            1
            if not successori
            else sum(cammini_verso_fine[s] for s in successori)
        )

    return cammini_da_start, cammini_verso_fine


@st.cache_data(show_spinner=False)
def analizza_da_dataframe(
    df: pd.DataFrame,
    ids_esclusi_per_data: frozenset[str] = frozenset(),
):
    grafo, anomalie = costruisci_grafo(df, set(ids_esclusi_per_data))
    if not nx.is_directed_acyclic_graph(grafo):
        cicli = list(nx.simple_cycles(grafo))[:20]
        return grafo, pd.DataFrame(), anomalie, cicli

    cammini_in, cammini_out = conta_cammini_dag(grafo)
    numero_nodi = grafo.number_of_nodes()
    if numero_nodi <= 1200:
        betweenness = nx.betweenness_centrality(grafo, normalized=True)
    else:
        betweenness = nx.betweenness_centrality(
            grafo, k=min(300, numero_nodi), normalized=True, seed=42
        )

    righe = []
    for nodo, attributi in grafo.nodes(data=True):
        fan_in = grafo.in_degree(nodo)
        fan_out = grafo.out_degree(nodo)
        percorsi = cammini_in[nodo] * cammini_out[nodo]
        log_percorsi = math.log10(percorsi) if percorsi > 0 else 0.0
        score = (
            log_percorsi
            + max(fan_out - 1, 0) * 1.2
            + max(fan_in - 1, 0) * 0.8
            + betweenness[nodo] * 10
        )
        righe.append(
            {
                "ID": nodo,
                "Nome": attributi.get("nome", ""),
                "PrismaMacroActivity": attributi.get("macro", ""),
                "Inizio": attributi.get("inizio"),
                "Fine": attributi.get("fine"),
                "Fan_in": fan_in,
                "Fan_out": fan_out,
                "Biforcazione": fan_out > 1,
                "Convergenza": fan_in > 1,
                "Cammini_da_start": cammini_in[nodo],
                "Cammini_verso_fine": cammini_out[nodo],
                "Percorsi_passanti": percorsi,
                "Log10_percorsi": log_percorsi,
                "Betweenness": betweenness[nodo],
                "Complexity_score": score,
                "Isolated": fan_in == 0 and fan_out == 0,
            }
        )

    ranking = pd.DataFrame(righe).sort_values(
        ["Complexity_score", "Percorsi_passanti", "Fan_out"],
        ascending=False,
    )
    return grafo, ranking, anomalie, []


def conta_percorsi_tra_nodi_dag(grafo: nx.DiGraph, origine: str, destinazione: str) -> int:
    """Conta i percorsi origine-destinazione in un DAG senza enumerarli."""
    if origine not in grafo or destinazione not in grafo:
        return 0
    raggiungibili = nx.descendants(grafo, origine) | {origine}
    antenati = nx.ancestors(grafo, destinazione) | {destinazione}
    utili = raggiungibili & antenati
    if destinazione not in utili:
        return 0
    sotto = grafo.subgraph(utili)
    conteggi = {n: 0 for n in sotto.nodes}
    conteggi[origine] = 1
    for nodo in nx.topological_sort(sotto):
        if nodo == origine:
            continue
        conteggi[nodo] = sum(conteggi[p] for p in sotto.predecessors(nodo))
    return int(conteggi.get(destinazione, 0))


def analizza_legami_ridondanti(grafo: nx.DiGraph) -> pd.DataFrame:
    """
    Individua archi diretti per i quali esiste anche un percorso indiretto.
    La segnalazione è strutturale: non implica automaticamente che il link possa
    essere cancellato, soprattutto in presenza di lag o relazioni diverse da FI.
    """
    colonne = [
        "Predecessor", "Predecessor name", "Successor", "Successor name",
        "Relationship", "Lag_g", "Alternative path", "Alternative steps",
        "Number of alternative paths", "Log10 percorsi alternativi",
        "Paths removed by deleting the link", "Path reduction pct",
        "Status", "Rationale",
    ]
    if not grafo or not nx.is_directed_acyclic_graph(grafo):
        return pd.DataFrame(columns=colonne)

    ridotto = nx.transitive_reduction(grafo)
    ridondanti = sorted(set(grafo.edges()) - set(ridotto.edges()))
    if not ridondanti:
        return pd.DataFrame(columns=colonne)

    cammini_da_origini, cammini_verso_fini = conta_cammini_dag(grafo)
    fini = [n for n in grafo if grafo.out_degree(n) == 0]
    percorsi_totali = sum(cammini_da_origini[n] for n in fini)
    righe = []

    for origine, destinazione in ridondanti:
        dati = grafo.edges[origine, destinazione]
        relazione = str(dati.get("relazione", "FI")).upper()
        lag = float(dati.get("lag", 0) or 0)
        test = grafo.copy()
        test.remove_edge(origine, destinazione)

        try:
            percorso = nx.shortest_path(test, origine, destinazione)
        except nx.NetworkXNoPath:
            continue

        alternativi = conta_percorsi_tra_nodi_dag(test, origine, destinazione)
        eliminati = int(cammini_da_origini[origine] * cammini_verso_fini[destinazione])
        riduzione_pct = eliminati / percorsi_totali * 100 if percorsi_totali else 0.0

        if relazione in {"FI", "FS"} and abs(lag) < 1e-9:
            stato = "High-priority review"
            motivazione = "FS with no lag and reachability already guaranteed by at least one indirect path"
        elif abs(lag) >= 1e-9:
            stato = "Review timing constraint"
            motivazione = "The direct link has a lag and may impose a specific timing constraint"
        else:
            stato = "Review relationship type"
            motivazione = f"Relationship {relazione}: redundant for reachability, but not necessarily for scheduling"

        righe.append({
            "Predecessor": origine,
            "Predecessor name": grafo.nodes[origine].get("nome", ""),
            "Successor": destinazione,
            "Successor name": grafo.nodes[destinazione].get("nome", ""),
            "Relationship": relazione,
            "Lag_g": lag,
            "Alternative path": " → ".join(map(str, percorso)),
            "Alternative steps": len(percorso) - 1,
            "Number of alternative paths": alternativi,
            "Log10 percorsi alternativi": math.log10(alternativi) if alternativi > 0 else 0.0,
            "Paths removed by deleting the link": eliminati,
            "Path reduction pct": riduzione_pct,
            "Status": stato,
            "Rationale": motivazione,
        })

    return pd.DataFrame(righe, columns=colonne).sort_values(
        ["Path reduction pct", "Number of alternative paths"],
        ascending=False,
    )


def analizza_grafo_esistente(grafo: nx.DiGraph) -> pd.DataFrame:
    """Ricalcola tutti gli indicatori su un grafo già costruito o modificato."""
    if not nx.is_directed_acyclic_graph(grafo):
        raise ValueError("The modified graph contains cycles.")
    cammini_in, cammini_out = conta_cammini_dag(grafo)
    numero_nodi = grafo.number_of_nodes()
    if numero_nodi <= 1200:
        betweenness = nx.betweenness_centrality(grafo, normalized=True)
    else:
        betweenness = nx.betweenness_centrality(
            grafo, k=min(300, numero_nodi), normalized=True, seed=42
        )
    righe = []
    for nodo, attributi in grafo.nodes(data=True):
        fan_in = grafo.in_degree(nodo)
        fan_out = grafo.out_degree(nodo)
        percorsi = cammini_in[nodo] * cammini_out[nodo]
        log_percorsi = math.log10(percorsi) if percorsi > 0 else 0.0
        score = (
            log_percorsi
            + max(fan_out - 1, 0) * 1.2
            + max(fan_in - 1, 0) * 0.8
            + betweenness[nodo] * 10
        )
        righe.append({
            "ID": nodo,
            "Nome": attributi.get("nome", ""),
            "PrismaMacroActivity": attributi.get("macro", ""),
            "Inizio": attributi.get("inizio"),
            "Fine": attributi.get("fine"),
            "Fan_in": fan_in,
            "Fan_out": fan_out,
            "Biforcazione": fan_out > 1,
            "Convergenza": fan_in > 1,
            "Cammini_da_start": cammini_in[nodo],
            "Cammini_verso_fine": cammini_out[nodo],
            "Percorsi_passanti": percorsi,
            "Log10_percorsi": log_percorsi,
            "Betweenness": betweenness[nodo],
            "Complexity_score": score,
            "Isolated": fan_in == 0 and fan_out == 0,
        })
    return pd.DataFrame(righe).sort_values(
        ["Complexity_score", "Percorsi_passanti", "Fan_out"],
        ascending=False,
    )


def crea_scenario_senza_ridondanze(
    grafo: nx.DiGraph,
    ridondanze: pd.DataFrame,
    modalita: str,
) -> tuple[nx.DiGraph, pd.DataFrame]:
    """
    Crea uno scenario rimuovendo le ridondanze selezionate.

    - Alta priorità: solo FI/FS senza lag.
    - Tutte: ogni ridondanza strutturale, indipendentemente da relazione e lag.
    """
    nuovo = grafo.copy()
    if ridondanze.empty:
        return nuovo, pd.DataFrame()
    if modalita == "All structural redundancies":
        selezionati = ridondanze.copy()
    else:
        selezionati = ridondanze[
            ridondanze["Status"] == "High-priority review"
        ].copy()
    rimossi = []
    for riga in selezionati.itertuples(index=False):
        origine = str(riga.Predecessor)
        destinazione = str(riga.Successor)
        if nuovo.has_edge(origine, destinazione):
            dati = dict(nuovo.edges[origine, destinazione])
            nuovo.remove_edge(origine, destinazione)
            rimossi.append({
                "Predecessor": origine,
                "Predecessor name": nuovo.nodes[origine].get("nome", ""),
                "Successor": destinazione,
                "Successor name": nuovo.nodes[destinazione].get("nome", ""),
                "Relationship": dati.get("relazione", "FI"),
                "Lag days": dati.get("lag", 0),
                "Original status": getattr(riga, "Status", ""),
                "Rationale": (
                    "Removal of every structural redundancy, regardless of relationship and lag"
                    if modalita == "All structural redundancies"
                    else "High-priority review: FS with no lag and an alternative indirect path"
                ),
            })
    return nuovo, pd.DataFrame(rimossi)


def ricostruisci_testo_predecessore(predecessore: str, relazione: str, lag: float) -> str:
    """Ricostruisce il formato predecessore accettato dall'import Excel dell'app."""
    relazione = str(relazione or "FI").upper()
    lag = float(lag or 0)
    lag_testo = ""
    if abs(lag) >= 1e-9:
        valore = int(lag) if lag.is_integer() else round(lag, 2)
        lag_testo = f"{'+' if lag > 0 else ''}{valore} g"
    if relazione in {"FI", "FS"} and not lag_testo:
        return str(predecessore)
    relazione_excel = "FI" if relazione == "FS" else relazione
    return f"{predecessore}{relazione_excel}{lag_testo}"


def crea_dataframe_gantt_ridotto(
    df_perimetro: pd.DataFrame,
    legami_rimossi: pd.DataFrame,
) -> pd.DataFrame:
    """Crea le sette colonne ricaricabili rimuovendo solo i link selezionati."""
    risultato = df_perimetro[
        [COL_MACRO, COL_ID, COL_PREDECESSORI, COL_NOME, COL_DURATA, COL_INIZIO, COL_FINE]
    ].copy()
    risultato[COL_ID] = risultato[COL_ID].apply(normalizza_id)
    archi_da_rimuovere = set()
    if not legami_rimossi.empty:
        archi_da_rimuovere = {
            (str(r.Predecessor), str(r.Successor))
            for r in legami_rimossi.itertuples(index=False)
        }

    nuovi_predecessori = []
    for _, riga in risultato.iterrows():
        successore = str(riga[COL_ID])
        parsed, non_riconosciuti = estrai_predecessori(riga[COL_PREDECESSORI])
        mantenuti = []
        for pred in parsed:
            origine = str(pred["id"])
            if (origine, successore) in archi_da_rimuovere:
                continue
            mantenuti.append(
                ricostruisci_testo_predecessore(
                    origine, pred["relazione"], float(pred["lag"] or 0)
                )
            )
        mantenuti.extend(non_riconosciuti)
        nuovi_predecessori.append(";".join(mantenuti))
    risultato[COL_PREDECESSORI] = nuovi_predecessori
    return risultato


def crea_excel_gantt_ridotto_ricaricabile(
    df_gantt_ridotto: pd.DataFrame,
    legami_rimossi: pd.DataFrame,
    confronto: pd.DataFrame,
) -> bytes:
    """Genera un workbook con foglio 'data' nuovamente caricabile nell'app."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_gantt_ridotto.to_excel(writer, sheet_name="data", index=False)
        legami_rimossi.to_excel(writer, sheet_name="Removed links", index=False)
        confronto.to_excel(writer, sheet_name="Comparison", index=False)
        ws = writer.sheets["data"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col, width in {"A": 32, "B": 10, "C": 48, "D": 60, "E": 18, "F": 20, "G": 20}.items():
            ws.column_dimensions[col].width = width
        for cell in ws["F"][1:]:
            cell.number_format = "dd/mm/yyyy hh:mm"
        for cell in ws["G"][1:]:
            cell.number_format = "dd/mm/yyyy hh:mm"
        for nome in ["Removed links", "Comparison"]:
            foglio = writer.sheets[nome]
            foglio.freeze_panes = "A2"
            foglio.auto_filter.ref = foglio.dimensions
            for col in foglio.columns:
                width = min(max(len(str(c.value)) if c.value is not None else 0 for c in col) + 2, 60)
                foglio.column_dimensions[col[0].column_letter].width = width
    output.seek(0)
    return output.getvalue()


def crea_excel_scenario_ridotto(
    ranking: pd.DataFrame,
    grafo: nx.DiGraph,
    legami_rimossi: pd.DataFrame,
    confronto: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        rimuovi_colonne_logaritmiche(ranking).to_excel(
            writer, sheet_name="Reduced scenario ranking", index=False
        )
        nx.to_pandas_edgelist(grafo).to_excel(
            writer, sheet_name="Remaining dependencies", index=False
        )
        legami_rimossi.to_excel(writer, sheet_name="Removed links", index=False)
        confronto.to_excel(writer, sheet_name="Comparison", index=False)
    output.seek(0)
    return output.getvalue()


def aggiungi_indicatori_ridondanza_ranking(
    ranking: pd.DataFrame, ridondanze: pd.DataFrame
) -> pd.DataFrame:
    risultato = ranking.copy()
    if ridondanze.empty:
        risultato["Incoming redundant links"] = 0
        risultato["Outgoing redundant links"] = 0
        return risultato
    entranti = ridondanze.groupby("Successor").size()
    uscenti = ridondanze.groupby("Predecessor").size()
    risultato["Incoming redundant links"] = risultato["ID"].map(entranti).fillna(0).astype(int)
    risultato["Outgoing redundant links"] = risultato["ID"].map(uscenti).fillna(0).astype(int)
    return risultato


def crea_grafo_ridondanza(
    grafo: nx.DiGraph, origine: str, destinazione: str
) -> go.Figure:
    """Evidenzia il link diretto in rosso e un percorso alternativo in blu."""
    test = grafo.copy()
    test.remove_edge(origine, destinazione)
    percorso = nx.shortest_path(test, origine, destinazione)
    archi_percorso = set(zip(percorso[:-1], percorso[1:]))
    nodi = set(percorso) | {origine, destinazione}
    for n in list(nodi):
        nodi.update(grafo.predecessors(n))
        nodi.update(grafo.successors(n))
    sotto = grafo.subgraph(nodi).copy()
    try:
        generazioni = list(nx.topological_generations(sotto))
        pos = {}
        for x, gen in enumerate(generazioni):
            ordinati = sorted(gen, key=str)
            centro = (len(ordinati) - 1) / 2
            for i, n in enumerate(ordinati):
                pos[n] = (x, centro - i)
    except Exception:
        pos = nx.spring_layout(sotto, seed=42)

    fig = go.Figure()
    gruppi = [
        ("Context", "#cbd5e1", "solid", [(u, v) for u, v in sotto.edges if (u, v) not in archi_percorso and (u, v) != (origine, destinazione)]),
        ("Indirect path", "#2563eb", "solid", list(archi_percorso)),
        ("Suspicious direct link", "#dc2626", "dash", [(origine, destinazione)]),
    ]
    for nome, colore, dash, edges in gruppi:
        xs, ys = [], []
        for u, v in edges:
            if u not in pos or v not in pos:
                continue
            xs += [pos[u][0], pos[v][0], None]
            ys += [pos[u][1], pos[v][1], None]
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="lines", name=nome, hoverinfo="skip",
            line=dict(color=colore, width=4 if nome != "Context" else 1, dash=dash),
        ))
    node_ids = list(sotto.nodes)
    colors = ["#dc2626" if n in {origine, destinazione} else "#2563eb" if n in percorso else "#94a3b8" for n in node_ids]
    fig.add_trace(go.Scatter(
        x=[pos[n][0] for n in node_ids], y=[pos[n][1] for n in node_ids],
        mode="markers+text", text=[str(n) for n in node_ids], textposition="middle center",
        hovertext=[f"{n} - {grafo.nodes[n].get('nome', '')}" for n in node_ids], hoverinfo="text",
        marker=dict(size=28, color=colors, line=dict(color="white", width=1)), name="Tasks",
    ))
    fig.update_layout(
        height=600, plot_bgcolor="white", margin=dict(l=10, r=10, t=45, b=60),
        title="Direct link vs. indirect path",
        xaxis=dict(visible=False), yaxis=dict(visible=False),
        legend=dict(orientation="h", y=-0.08),
    )
    return fig


def crea_excel_ridondanze(ridondanze: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        rimuovi_colonne_logaritmiche(ridondanze).to_excel(
            writer, sheet_name="Redundant links", index=False
        )
        ws = writer.sheets["Redundant links"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            width = min(max(len(str(c.value)) if c.value is not None else 0 for c in col) + 2, 60)
            ws.column_dimensions[col[0].column_letter].width = width
        headers = {c.value: c.column for c in ws[1]}
        for name in ["Number of alternative paths", "Paths removed by deleting the link"]:
            if name in headers:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row, headers[name]).number_format = "#,##0"
        if "Path reduction pct" in headers:
            for row in range(2, ws.max_row + 1):
                ws.cell(row, headers["Path reduction pct"]).number_format = '0.000"%"'
    output.seek(0)
    return output.getvalue()


def aggiungi_posizione_temporale(
    ranking: pd.DataFrame,
    data_riferimento: pd.Timestamp,
    data_fine_periodo: pd.Timestamp,
) -> pd.DataFrame:
    """Posiziona la fine attività tra data riferimento (0%) e fine periodo (100%)."""
    risultato = ranking.copy()
    durata_periodo = (data_fine_periodo.normalize() - data_riferimento.normalize()).days
    if durata_periodo <= 0:
        risultato["Posizione_temporale_pct"] = 0.0
        return risultato

    risultato["Posizione_temporale_pct"] = (
        (risultato["Fine"].dt.normalize() - data_riferimento.normalize()).dt.days
        / durata_periodo
    ).mul(100).clip(lower=0, upper=100)
    return risultato


def nodi_da_visualizzare(
    grafo: nx.DiGraph,
    ranking_filtrato: pd.DataFrame,
    top_n: int,
    raggio: int,
) -> set[str]:
    nodi_top = set(ranking_filtrato.head(top_n)["ID"])
    nodi = set(nodi_top)
    inverso = grafo.reverse(copy=False)

    for nodo in nodi_top:
        nodi.update(nx.single_source_shortest_path_length(grafo, nodo, cutoff=raggio))
        nodi.update(nx.single_source_shortest_path_length(inverso, nodo, cutoff=raggio))
    return nodi


def nodi_sui_percorsi_verso_attivita(
    grafo: nx.DiGraph, attivita: str
) -> set[str]:
    """Restituisce l'attività selezionata e tutti i suoi antenati."""
    if attivita not in grafo:
        return set()
    return set(nx.ancestors(grafo, attivita)) | {attivita}


def crea_grafo_interattivo(
    grafo: nx.DiGraph,
    ranking: pd.DataFrame,
    nodi: set[str],
    mostra_etichette: bool,
    layout_grafo: str,
) -> go.Figure:
    sotto_grafo = grafo.subgraph(nodi).copy()
    if sotto_grafo.number_of_nodes() == 0:
        return go.Figure()

    # Layout selezionabile. Il layout topologico evidenzia il flusso temporale,
    # mentre gli altri aiutano a leggere cluster, centralità e densità dei legami.
    if layout_grafo == "Topologico (sinistra-destra)":
        try:
            generazioni = list(nx.topological_generations(sotto_grafo))
            pos = {}
            for x, generazione in enumerate(generazioni):
                ordinati = sorted(
                    generazione,
                    key=lambda n: (
                        str(sotto_grafo.nodes[n].get("macro", "")),
                        str(n),
                    ),
                )
                centro = (len(ordinati) - 1) / 2
                for i, nodo in enumerate(ordinati):
                    pos[nodo] = (x, centro - i)
        except nx.NetworkXUnfeasible:
            pos = nx.spring_layout(sotto_grafo, seed=42, k=1.5)
    elif layout_grafo == "Forze (Spring)":
        pos = nx.spring_layout(sotto_grafo, seed=42, k=1.5, iterations=100)
    elif layout_grafo == "Kamada-Kawai":
        pos = nx.kamada_kawai_layout(sotto_grafo)
    elif layout_grafo == "Circolare":
        pos = nx.circular_layout(sotto_grafo)
    elif layout_grafo == "Spettrale":
        pos = nx.spectral_layout(sotto_grafo)
    else:
        pos = nx.spring_layout(sotto_grafo, seed=42, k=1.5)

    figura = go.Figure()
    relazioni_presenti = sorted(
        {d.get("relazione", "FI") for _, _, d in sotto_grafo.edges(data=True)}
    )

    for relazione in relazioni_presenti:
        x_edge, y_edge = [], []
        for origine, destinazione, dati in sotto_grafo.edges(data=True):
            if dati.get("relazione", "FI") != relazione:
                continue
            x0, y0 = pos[origine]
            x1, y1 = pos[destinazione]
            x_edge.extend([x0, x1, None])
            y_edge.extend([y0, y1, None])
        figura.add_trace(
            go.Scatter(
                x=x_edge,
                y=y_edge,
                mode="lines",
                line=dict(color=RELATION_COLORS.get(relazione, "#94a3b8"), width=1.2),
                hoverinfo="skip",
                name={"FI": "FS", "II": "SS", "IF": "SF"}.get(relazione, relazione),
            )
        )

    ranking_map = ranking.set_index("ID")
    node_ids = list(sotto_grafo.nodes())
    scores = [float(ranking_map.loc[n, "Complexity_score"]) for n in node_ids]
    dimensioni = [14 + min(30, 6 * math.log1p(max(score, 0))) for score in scores]

    hover = []
    for nodo in node_ids:
        riga = ranking_map.loc[nodo]
        hover.append(
            f"<b>{nodo} - {riga['Nome']}</b><br>"
            f"Macro: {riga['PrismaMacroActivity']}<br>"
            f"Predecessors: {riga['Fan_in']}<br>"
            f"Successors: {riga['Fan_out']}<br>"
            f"Betweenness: {riga['Betweenness']:.4f}<br>"
            f"Score: {riga['Complexity_score']:.2f}"
        )

    figura.add_trace(
        go.Scatter(
            x=[pos[n][0] for n in node_ids],
            y=[pos[n][1] for n in node_ids],
            mode="markers+text" if mostra_etichette else "markers",
            text=[str(n) for n in node_ids] if mostra_etichette else None,
            textposition="middle center",
            hovertext=hover,
            hoverinfo="text",
            marker=dict(
                size=dimensioni,
                color=scores,
                colorscale="Turbo",
                showscale=True,
                colorbar=dict(title="Complexity"),
                line=dict(width=1, color="white"),
            ),
            name="Tasks",
        )
    )

    figura.update_layout(
        height=800,
        margin=dict(l=10, r=10, t=75, b=80),
        title=dict(
            text="Interactive dependency graph",
            x=0.01,
            xanchor="left",
            y=0.98,
            yanchor="top",
        ),
        plot_bgcolor="white",
        hovermode="closest",
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
        # Legenda spostata sotto il grafico per evitare sovrapposizioni col titolo.
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.06,
            xanchor="left",
            x=0,
        ),
    )
    return figura


def rimuovi_colonne_logaritmiche(df: pd.DataFrame) -> pd.DataFrame:
    """Esclude da tabelle ed export tutte le variabili logaritmiche."""
    colonne_da_togliere = [
        col for col in df.columns
        if "log10" in str(col).lower() or "logarit" in str(col).lower()
    ]
    return df.drop(columns=colonne_da_togliere, errors="ignore").copy()


def crea_excel_ranking(ranking_visualizzato: pd.DataFrame) -> bytes:
    """Crea un Excel contenente esattamente il ranking mostrato a video."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        rimuovi_colonne_logaritmiche(ranking_visualizzato).to_excel(
            writer, sheet_name="Complexity ranking", index=False
        )
        worksheet = writer.sheets["Complexity ranking"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        formati_interi = {
            "Fan_in",
            "Fan_out",
            "Cammini_da_start",
            "Cammini_verso_fine",
            "Percorsi_passanti",
        }
        intestazioni = {
            cell.value: cell.column for cell in worksheet[1] if cell.value is not None
        }
        for nome_colonna in formati_interi:
            indice_colonna = intestazioni.get(nome_colonna)
            if indice_colonna is not None:
                for riga in range(2, worksheet.max_row + 1):
                    worksheet.cell(riga, indice_colonna).number_format = "#,##0"

        for nome_colonna, formato in {
            "Betweenness": "0.0000",
            "Complexity_score": "0.00",
            "Posizione_temporale_pct": '0.0"%"',
        }.items():
            indice_colonna = intestazioni.get(nome_colonna)
            if indice_colonna is not None:
                for riga in range(2, worksheet.max_row + 1):
                    worksheet.cell(riga, indice_colonna).number_format = formato

        for colonna in worksheet.columns:
            larghezza = min(
                max(len(str(cella.value)) if cella.value is not None else 0 for cella in colonna) + 2,
                45,
            )
            worksheet.column_dimensions[colonna[0].column_letter].width = larghezza

    output.seek(0)
    return output.getvalue()


def crea_excel_output(
    ranking: pd.DataFrame,
    anomalie: pd.DataFrame,
    grafo: nx.DiGraph,
    ridondanze: pd.DataFrame | None = None,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        rimuovi_colonne_logaritmiche(ranking).to_excel(
            writer, sheet_name="Ranking", index=False
        )
        anomalie.to_excel(writer, sheet_name="Issues", index=False)
        nx.to_pandas_edgelist(grafo).to_excel(
            writer, sheet_name="Dependencies", index=False
        )
        if ridondanze is not None:
            rimuovi_colonne_logaritmiche(ridondanze).to_excel(
                writer, sheet_name="Redundant links", index=False
            )
    output.seek(0)
    return output.getvalue()


def main() -> None:
    st.set_page_config(
        page_title="Gantt Complexity Analyzer",
        page_icon="🕸️",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.title("Gantt Complexity Analyzer")
    st.caption(
        "Upload a Microsoft Project (.mpp) or Excel schedule directly, "
        "explore the dependency graph, and identify nodes that multiply paths."
    )

    file_schedule = st.sidebar.file_uploader(
        "Upload schedule",
        type=["mpp", "xlsx", "xlsm"],
        help="MPP files are read directly without opening Microsoft Project.",
    )

    if file_schedule is None:
        st.info("Upload an .mpp, .xlsx, or .xlsm file from the sidebar to begin.")
        st.stop()

    contenuto = file_schedule.getvalue()
    estensione = Path(file_schedule.name).suffix.lower()

    try:
        if estensione == ".mpp":
            with st.spinner("Reading Microsoft Project file directly..."):
                df_completo = carica_schedule_mpp(contenuto, file_schedule.name)
            st.sidebar.success(
                f"MPP read directly: {len(df_completo)} tasks extracted."
            )
            st.sidebar.download_button(
                "Download MPP extraction as Excel",
                data=crea_excel_schedule_estratto(df_completo),
                file_name="schedule_extracted_from_mpp.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            fogli = leggi_fogli(contenuto)
            indice_default = fogli.index("data") if "data" in fogli else 0
            foglio = st.sidebar.selectbox("Excel worksheet", fogli, index=indice_default)
            df_completo = carica_schedule(contenuto, foglio)
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    # Le summary sono già state escluse dai loader e non entrano mai nell'analisi.
    numero_summary_escluse = int(df_completo.attrs.get("summary_escluse", 0))
    ids_summary_esclusi = set()
    if df_completo.empty:
        st.warning("No operational tasks are available in the uploaded file.")
        st.stop()
    st.sidebar.subheader("1. Analysis scope")
    st.sidebar.caption(f"Summary tasks ignored during upload: {numero_summary_escluse}")
    # Per ogni nuovo file la reference date parte dalla prima data valida
    # del Gantt. L'utente può modificarla o riportarla rapidamente a oggi.
    date_inizio_valide = df_completo[COL_INIZIO].dropna()
    if date_inizio_valide.empty:
        st.error("The uploaded schedule contains no valid start dates.")
        st.stop()
    prima_data_gantt = pd.Timestamp(date_inizio_valide.min()).normalize().date()
    firma_file = f"{file_schedule.name}:{len(contenuto)}:{hash(contenuto)}"
    if st.session_state.get("firma_file_data_riferimento") != firma_file:
        st.session_state["data_riferimento_gantt"] = prima_data_gantt
        st.session_state["firma_file_data_riferimento"] = firma_file

    if st.sidebar.button("Set today's date", use_container_width=True):
        st.session_state["data_riferimento_gantt"] = pd.Timestamp.today().date()

    data_riferimento_input = st.sidebar.date_input(
        "Reference date",
        key="data_riferimento_gantt",
        help=(
            "By default, this is the first date in the schedule. "
            "Tasks with a finish date earlier than the selected date are excluded."
        ),
    )
    st.sidebar.caption(f"First schedule date: {prima_data_gantt:%d/%m/%Y}")
    data_riferimento = pd.Timestamp(data_riferimento_input).normalize()

    maschera_escluse_data = (
        df_completo[COL_FINE].notna()
        & (df_completo[COL_FINE].dt.normalize() < data_riferimento)
    )
    ids_esclusi_per_data = set(df_completo.loc[maschera_escluse_data, COL_ID])
    df_dopo_data = df_completo.loc[~maschera_escluse_data].copy()

    if df_dopo_data.empty:
        st.warning("No tasks remain after applying the reference date.")
        st.stop()

    # Il grafo preliminare serve esclusivamente a determinare quali attività
    # conducono all'eventuale attività finale. None metrica viene calcolata qui.
    grafo_preliminare, anomalie_preliminari = costruisci_grafo(
        df_dopo_data, ids_esclusi_per_data | ids_summary_esclusi
    )

    usa_attivita_finale = st.sidebar.checkbox(
        "Select a final task",
        value=False,
        help="When enabled, only the selected task and all tasks that can reach it are retained.",
    )

    attivita_selezionata = None
    if usa_attivita_finale:
        opzioni_finali = df_dopo_data.sort_values(
            [COL_FINE, COL_NOME, COL_ID],
            ascending=[False, True, True],
            na_position="last",
        )[[COL_ID, COL_NOME, COL_FINE]]
        etichette_finali = {
            (
                f"{r.ID} - {r.Nome} | "
                f"End nodessh: {pd.Timestamp(r.Fine):%d/%m/%Y}"
                if pd.notna(r.Fine)
                else f"{r.ID} - {r.Nome} | End nodessh: not available"
            ): r.ID
            for r in opzioni_finali.itertuples(index=False)
        }
        etichetta_finale = st.sidebar.selectbox(
            "Final task",
            options=list(etichette_finali.keys()),
            index=None,
            placeholder="Select task ID or name",
        )
        if etichetta_finale is None:
            st.info("Select the final task in the sidebar.")
            st.stop()
        attivita_selezionata = etichette_finali[etichetta_finale]

        ids_percorso_finale = (
            set(nx.ancestors(grafo_preliminare, attivita_selezionata))
            | {attivita_selezionata}
        )
        df = df_dopo_data[df_dopo_data[COL_ID].isin(ids_percorso_finale)].copy()
        ids_esclusi_per_finale = set(df_dopo_data[COL_ID]) - ids_percorso_finale
    else:
        df = df_dopo_data.copy()
        ids_esclusi_per_finale = set()

    ids_esclusi_totali = frozenset(
        ids_summary_esclusi | ids_esclusi_per_data | ids_esclusi_per_finale
    )

    try:
        with st.spinner("Building and analyzing the filtered graph..."):
            grafo, ranking, anomalie, cicli = analizza_da_dataframe(
                df, ids_esclusi_totali
            )
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    if cicli:
        st.error("The filtered graph contains cycles. Fix the dependencies before calculating paths.")
        st.dataframe(
            pd.DataFrame({"Cycle": [" -> ".join(ciclo) for ciclo in cicli]}),
            use_container_width=True,
            hide_index=True,
        )
        st.stop()

    if usa_attivita_finale:
        fine_selezionata = df.loc[df[COL_ID] == attivita_selezionata, COL_FINE]
        if fine_selezionata.empty or pd.isna(fine_selezionata.iloc[0]):
            st.error("The selected final task does not have a valid finish date.")
            st.stop()
        data_fine_periodo = pd.Timestamp(fine_selezionata.iloc[0]).normalize()
        descrizione_fine_periodo = (
            f"final task {attivita_selezionata} - "
            f"{grafo.nodes[attivita_selezionata].get('nome', '')}"
        )
    else:
        date_fine_valide = ranking["Fine"].dropna()
        if date_fine_valide.empty:
            st.error("There are no valid finish dates for calculating the time position.")
            st.stop()
        data_fine_periodo = pd.Timestamp(date_fine_valide.max()).normalize()
        descrizione_fine_periodo = "latest task in the analyzed scope"

    ranking = aggiungi_posizione_temporale(
        ranking, data_riferimento, data_fine_periodo
    )
    with st.spinner("Analyzing potentially redundant links..."):
        ridondanze_originali = analizza_legami_ridondanti(grafo)

    st.sidebar.divider()
    st.sidebar.subheader("2. Scenario without redundancies")
    modalita_riduzione = st.sidebar.radio(
        "Redundancies to remove",
        [
            "None",
            "High priority only",
            "All structural redundancies",
        ],
        index=0,
        help=(
            "High priority removes only FS links with no lag. All structural "
            "redundancies also removes links with lag or SS, FF, and SF relationships. "
            "The source file is not modified."
        ),
    )
    usa_scenario_ridotto = modalita_riduzione != "None"

    grafo_originale = grafo.copy()
    ranking_originale = ranking.copy()
    legami_rimossi = pd.DataFrame()
    confronto_scenari = pd.DataFrame()
    df_gantt_ridotto_ricaricabile = pd.DataFrame()

    if usa_scenario_ridotto:
        grafo, legami_rimossi = crea_scenario_senza_ridondanze(
            grafo_originale, ridondanze_originali, modalita_riduzione
        )
        with st.spinner("Fully recalculating the scenario without redundancies..."):
            ranking = analizza_grafo_esistente(grafo)
            ranking = aggiungi_posizione_temporale(
                ranking, data_riferimento, data_fine_periodo
            )
            ridondanze = analizza_legami_ridondanti(grafo)

        tot_originale = sum(
            conta_cammini_dag(grafo_originale)[0][n]
            for n in grafo_originale
            if grafo_originale.out_degree(n) == 0
        )
        tot_ridotto = sum(
            conta_cammini_dag(grafo)[0][n]
            for n in grafo
            if grafo.out_degree(n) == 0
        )
        confronto_scenari = pd.DataFrame([
            {
                "Metric": "Dependencies",
                "Original scenario": grafo_originale.number_of_edges(),
                "Reduced scenario": grafo.number_of_edges(),
                "Change": grafo.number_of_edges() - grafo_originale.number_of_edges(),
            },
            {
                "Metric": "Total paths to end nodes",
                "Original scenario": tot_originale,
                "Reduced scenario": tot_ridotto,
                "Change": tot_ridotto - tot_originale,
            },
            {
                "Metric": "Remaining redundancies",
                "Original scenario": len(ridondanze_originali),
                "Reduced scenario": len(ridondanze),
                "Change": len(ridondanze) - len(ridondanze_originali),
            },
        ])
        df_gantt_ridotto_ricaricabile = crea_dataframe_gantt_ridotto(
            df, legami_rimossi
        )
        st.warning(
            f"Experimental scenario enabled ({modalita_riduzione}): removed "
            f"{len(legami_rimossi)} link. All i grafici e gli indicatori sottostanti "
            "sono ricalcolati. Il file sorgente non viene modificato."
        )
    else:
        ridondanze = ridondanze_originali

    ranking = aggiungi_indicatori_ridondanza_ranking(ranking, ridondanze)

    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    c1.metric("Analyzed tasks", grafo.number_of_nodes())
    c2.metric("Excluded by date", len(ids_esclusi_per_data))
    c3.metric("Excluded by final task", len(ids_esclusi_per_finale))
    c4.metric("Dependencies", grafo.number_of_edges())
    c5.metric("Start nodes", sum(grafo.in_degree(n) == 0 for n in grafo.nodes))
    c6.metric("End nodes", sum(grafo.out_degree(n) == 0 for n in grafo.nodes))
    c7.metric("Issues", len(anomalie))

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Excluded summary tasks", numero_summary_escluse)
    r2.metric("Potentially redundant links", len(ridondanze))
    r3.metric(
        "High-priority review",
        int((ridondanze["Status"] == "High-priority review").sum()) if not ridondanze.empty else 0,
    )
    r4.metric(
        "Affected tasks",
        len(set(ridondanze["Predecessor"]) | set(ridondanze["Successor"])) if not ridondanze.empty else 0,
    )
    if usa_attivita_finale:
        st.success(
            "All metrics are recalculated on the filtered scope: "
            f"reference date {data_riferimento:%d/%m/%Y} and single final task "
            f"{attivita_selezionata}."
        )

    ranking_visualizzato = ranking.sort_values(
        ["Cammini_da_start", "Percorsi_passanti", "ID"],
        ascending=[False, False, True],
    )

    tab_ranking, tab_ridondanze, tab_confronto, tab_grafo, tab_anomalie = st.tabs(
        [
            "Complexity ranking",
            "Redundant links",
            "Reduced scenario",
            "Interactive graph",
            "Issues",
        ]
    )

    with tab_grafo:
        st.subheader("Display controls")
        ctrl1, ctrl2, ctrl3 = st.columns(3)
        with ctrl1:
            mostra_tutto = st.checkbox(
                "Show 100% of tasks", value=False,
                help="When enabled, node count and context levels are ignored.",
                key="grafo_mostra_tutto",
            )
            mostra_etichette = st.checkbox(
                "Show IDs in nodes", value=True, key="grafo_etichette"
            )
        with ctrl2:
            top_n = st.slider(
                "Number of primary nodes", 5, 60, 20,
                disabled=mostra_tutto, key="grafo_top_n",
            )
        with ctrl3:
            raggio = st.slider(
                "Context levels", 0, 3, 1,
                disabled=mostra_tutto, key="grafo_raggio",
            )

        ranking_grafo = ranking.copy()
        if mostra_tutto:
            nodi = set(ranking_grafo["ID"])
        else:
            nodi = nodi_da_visualizzare(grafo, ranking_grafo, top_n, raggio)
        layout_grafo = "Topologico (sinistra-destra)"
        ids_base_visualizzazione = set(ranking_grafo["ID"])
        nodi_rappresentati = nodi & set(grafo.nodes)
        totale_base = len(ids_base_visualizzazione)
        pct_nodi = len(nodi_rappresentati) / totale_base * 100 if totale_base else 0.0
        archi_base = {
            (u, v) for u, v in grafo.edges
            if u in ids_base_visualizzazione and v in ids_base_visualizzazione
        }
        archi_rappresentati = {
            (u, v) for u, v in grafo.edges
            if u in nodi_rappresentati and v in nodi_rappresentati
        }
        pct_archi = len(archi_rappresentati) / len(archi_base) * 100 if archi_base else 0.0
        figura = crea_grafo_interattivo(
            grafo, ranking, nodi, mostra_etichette, layout_grafo
        )

        cop1, cop2, cop3 = st.columns(3)
        cop1.metric(
            "Task coverage",
            f"{pct_nodi:.1f}%",
            help="Percentage of tasks in the filtered scope included in the graph by the sliders.",
        )
        cop2.metric(
            "Displayed tasks",
            f"{len(nodi_rappresentati):,}".replace(",", "."),
            delta=f"su {totale_base:,}".replace(",", "."),
            delta_color="off",
        )
        cop3.metric(
            "Dependency coverage",
            f"{pct_archi:.1f}%",
            help="Percentage of links within the filtered scope shown in the view.",
        )
        st.progress(min(max(pct_nodi / 100.0, 0.0), 1.0))
        st.caption(
            "Coverage changes with Number of primary nodes and Context levels. "
            "When Show 100% of tasks is enabled, the sliders are ignored. "
            "The denominator includes all tasks in the current scope."
        )
        st.plotly_chart(figura, use_container_width=True)
        if usa_attivita_finale:
            st.success(
                f"The analyzed graph ends only at {attivita_selezionata} - "
                f"{grafo.nodes[attivita_selezionata].get('nome', '')}."
            )
        st.caption(
            "Hover over the nodes to view details. Color and size increase "
            "with the complexity score."
        )

    with tab_ranking:
        st.caption(
            f"Time position: 0% = {data_riferimento:%d/%m/%Y}; "
            f"100% = {data_fine_periodo:%d/%m/%Y} ({descrizione_fine_periodo})."
        )
        colonne = [
            "ID",
            "Nome",
            "Inizio",
            "Fine",
            "Posizione_temporale_pct",
            "Cammini_da_start",
            "Fan_in",
            "Fan_out",
            "Cammini_verso_fine",
            "Percorsi_passanti",
            "Betweenness",
            "Complexity_score",
            "Incoming redundant links",
            "Outgoing redundant links",
            "PrismaMacroActivity",
            "Isolated",
        ]
        nomi_colonne_ranking = {
            "Nome": "Name",
            "Inizio": "Start",
            "Fine": "Finish",
            "PrismaMacroActivity": "Prisma Macro Activity",
            "Posizione_temporale_pct": "Time position",
            "Fan_in": "Fan in",
            "Fan_out": "Fan out",
            "Cammini_da_start": "Paths from start",
            "Cammini_verso_fine": "Paths to end",
            "Percorsi_passanti": "Paths through task",
            "Complexity_score": "Complexity score",
            "Incoming redundant links": "Incoming redundant links",
            "Outgoing redundant links": "Outgoing redundant links",
        }
        ranking_tabella = ranking_visualizzato[colonne].rename(
            columns=nomi_colonne_ranking
        )
        st.dataframe(
            ranking_tabella,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Time position": st.column_config.ProgressColumn(
                    "Time position",
                    format="%.1f%%",
                    min_value=0.0,
                    max_value=100.0,
                ),
                "Fan in": st.column_config.NumberColumn(format="localized"),
                "Fan out": st.column_config.NumberColumn(format="localized"),
                "Paths from start": st.column_config.NumberColumn(format="localized"),
                "Paths to end": st.column_config.NumberColumn(format="localized"),
                "Paths through task": st.column_config.NumberColumn(format="localized"),
                "Betweenness": st.column_config.NumberColumn(format="%.4f"),
                "Complexity score": st.column_config.NumberColumn(format="%.2f"),
                "Incoming redundant links": st.column_config.NumberColumn(format="localized"),
                "Outgoing redundant links": st.column_config.NumberColumn(format="localized"),
                "Isolated": st.column_config.CheckboxColumn("Isolated"),
            },
        )
        st.download_button(
            "Download complexity ranking as Excel",
            data=crea_excel_ranking(ranking_tabella),
            file_name="displayed_complexity_ranking.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with tab_ridondanze:
        st.warning(
            "Structural redundancy does not automatically mean that a link is unnecessary. "
            "Links with lag or a relationship other than FS may impose specific timing constraints."
        )
        if ridondanze.empty:
            st.success("No direct links with an alternative indirect path were found in the analyzed scope.")
        else:
            filtro_stato = st.selectbox(
                "Status filter",
                ["All"] + sorted(ridondanze["Status"].unique().tolist()),
                key="filtro_ridondanze",
            )
            tabella_rid = ridondanze if filtro_stato == "All" else ridondanze[ridondanze["Status"] == filtro_stato]
            colonne_ridondanze = [
                "Predecessor",
                "Predecessor name",
                "Successor",
                "Successor name",
                "Relationship",
                "Lag_g",
                "Alternative path",
                "Paths removed by deleting the link",
                "Alternative steps",
                "Number of alternative paths",
                "Path reduction pct",
                "Status",
                "Rationale",
            ]
            nomi_ridondanze = {
                "Predecessor": "Predecessor",
                "Predecessor name": "Predecessor name",
                "Successor": "Successor",
                "Successor name": "Successor name",
                "Relationship": "Relationship",
                "Lag_g": "Lag days",
                "Alternative path": "Alternative path",
                "Paths removed by deleting the link": "Paths removed by deleting the link",
                "Alternative steps": "Alternative steps",
                "Number of alternative paths": "Number of alternative paths",
                "Path reduction pct": "Path reduction pct",
                "Status": "Status",
                "Rationale": "Rationale",
            }
            tabella_rid_visualizzata = tabella_rid[colonne_ridondanze].rename(
                columns=nomi_ridondanze
            )
            tabella_rid_visualizzata["Relationship"] = (
                tabella_rid_visualizzata["Relationship"]
                .replace({"FI": "FS", "II": "SS", "IF": "SF"})
            )
            st.dataframe(
                tabella_rid_visualizzata,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Number of alternative paths": st.column_config.NumberColumn(format="localized"),
                    "Paths removed by deleting the link": st.column_config.NumberColumn(format="localized"),
                    "Path reduction pct": st.column_config.NumberColumn(format="%.3f%%"),
                    "Lag days": st.column_config.NumberColumn(format="%.2f"),
                },
            )
            st.download_button(
                "Download redundant links as Excel",
                data=crea_excel_ridondanze(tabella_rid_visualizzata),
                file_name="potentially_redundant_links.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            opzioni_rid = {
                f"{r.Predecessor} → {r.Successor} | {r.Status}": (r.Predecessor, r.Successor)
                for r in tabella_rid.itertuples(index=False)
            }
            selezione_rid = st.selectbox(
                "Select a link to compare it with the indirect path",
                list(opzioni_rid.keys()),
                index=None,
                placeholder="Select predecessor → successor",
            )
            if selezione_rid:
                origine_rid, destinazione_rid = opzioni_rid[selezione_rid]
                st.plotly_chart(
                    crea_grafo_ridondanza(grafo, origine_rid, destinazione_rid),
                    use_container_width=True,
                )
                dettaglio = ridondanze[
                    (ridondanze["Predecessor"] == origine_rid)
                    & (ridondanze["Successor"] == destinazione_rid)
                ].iloc[0]
                st.info(
                    f"Displayed alternative path: {dettaglio['Alternative path']}. "
                    f"Total indirect paths between the two tasks: "
                    f"{int(dettaglio['Number of alternative paths']):,}.".replace(",", ".")
                )

    with tab_confronto:
        if not usa_scenario_ridotto:
            st.info(
                "Select an option in the 'Scenario without redundancies' section "
                "to generate and analyze the reduced scenario."
            )
        elif legami_rimossi.empty:
            st.success("No links need to be removed for the selected option.")
        else:
            if modalita_riduzione == "All structural redundancies":
                st.error(
                    "Extreme scenario: links with lag or relationships other than FS were also "
                    "removed. Reachability remains covered by indirect paths, but timing constraints "
                    "may change. Validate each removal with the planner."
                )
            else:
                st.warning(
                    "The reduced scenario is a review tool. Validate each link with the planner "
                    "before modifying the official schedule."
                )
            st.subheader("Original vs. reduced scenario")
            st.dataframe(confronto_scenari, use_container_width=True, hide_index=True)
            st.subheader("Links removed in the scenario")
            st.dataframe(legami_rimossi, use_container_width=True, hide_index=True)
            st.download_button(
                "Download reloadable reduced schedule",
                data=crea_excel_gantt_ridotto_ricaricabile(
                    df_gantt_ridotto_ricaricabile,
                    legami_rimossi,
                    confronto_scenari,
                ),
                file_name=(
                    "reloadable_schedule_without_all_redundancies.xlsx"
                    if modalita_riduzione == "All structural redundancies"
                    else "reloadable_schedule_without_high_priority_redundancies.xlsx"
                ),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="The data worksheet can be uploaded again into the app as an Excel schedule.",
            )
            st.caption(
                "The 'data' worksheet retains the PrismaMacroActivity, ID, Predecessori, Nome, "
                "Durata, Inizio, and Fine column names for reload compatibility. Predecessors are "
                f"removed according to the selected option: {modalita_riduzione}."
            )
            st.download_button(
                "Download reduced scenario analysis",
                data=crea_excel_scenario_ridotto(
                    ranking, grafo, legami_rimossi, confronto_scenari
                ),
                file_name=(
                    "analysis_without_all_redundancies.xlsx"
                    if modalita_riduzione == "All structural redundancies"
                    else "analysis_without_high_priority_redundancies.xlsx"
                ),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with tab_anomalie:
        if anomalie.empty:
            st.success("No missing or unrecognized predecessors.")
        else:
            st.dataframe(anomalie, use_container_width=True, hide_index=True)

    st.sidebar.divider()
    st.sidebar.download_button(
        "Download Excel analysis",
        data=crea_excel_output(ranking, anomalie, grafo, ridondanze),
        file_name=("reduced_scenario_analysis.xlsx" if usa_scenario_ridotto else "complexity_ranking.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.sidebar.download_button(
        "Download ranking CSV",
        data=rimuovi_colonne_logaritmiche(ranking).to_csv(index=False).encode("utf-8-sig"),
        file_name="complexity_ranking.csv",
        mime="text/csv",
    )

    with st.expander("How complexity is calculated"):
        st.markdown(
            """
- **Time position**: percentage position of the task finish between the reference date (0%) and the end of the period (100%). If a final task is selected, its finish date dynamically becomes 100%.
- **Fan-in**: number of predecessors converging on the task.
- **Fan-out**: number of successors generated by the task.
- **Paths through task**: paths from the filtered start nodes to the task multiplied by paths from the task to the end nodes. If a final task is selected, it is the graph's only end node.
- **Betweenness**: importance of the node as a bridge in the network.
- **Complexity score**: logarithmic combination of paths, branching, convergence, and centrality.
- **Potentially redundant link**: direct link between two tasks for which at least one indirect path also exists. The planner should review the finding, especially when the link has lag or a relationship other than FS.

The score measures the structural complexity of the network and does not replace critical-path calculation.
            """
        )


if __name__ == "__main__":
    main()
