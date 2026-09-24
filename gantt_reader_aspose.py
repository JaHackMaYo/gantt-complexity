from __future__ import annotations

import tempfile
from collections import defaultdict
from pathlib import Path

import pandas as pd

COL_MACRO = "PrismaMacroActivity"
COL_ID = "ID"
COL_UNIQUE_ID = "Unique ID"
COL_PREDECESSORI = "Predecessori"
COL_NOME = "Nome"
COL_DURATA = "Durata"
COL_INIZIO = "Inizio"
COL_FINE = "Fine"


def formatta_lag(giorni: float) -> str:
    if abs(giorni) < 1e-9:
        return ""
    valore = int(giorni) if float(giorni).is_integer() else round(giorni, 2)
    return f"{'+' if giorni > 0 else ''}{valore} g"

READER_NAME = "Aspose.Tasks"
REQUIREMENTS = "aspose-tasks"

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


def costruisci_predecessori(project) -> tuple[dict[int, list[str]], list[dict]]:
    predecessori = defaultdict(list)
    dettaglio = []
    archi_visti = set()
    for link in project.task_links:
        pred = link.pred_task
        succ = link.succ_task
        if pred is None or succ is None:
            continue
        pred_id = int(pred.id)
        succ_id = int(succ.id)
        if pred_id == succ_id:
            continue
        relazione = codice_relazione(link.link_type)
        lag_giorni = numero_lag_giorni(link)
        chiave = (pred_id, succ_id, relazione, round(lag_giorni, 8))
        if chiave in archi_visti:
            continue
        archi_visti.add(chiave)
        lag_testo = formatta_lag(lag_giorni)
        testo = str(pred_id) if relazione == "FI" and not lag_testo else f"{pred_id}{relazione}{lag_testo}"
        predecessori[succ_id].append(testo)
        dettaglio.append({
            "Predecessore": str(pred_id),
            "Predecessore Unique ID": str(getattr(pred, "uid", "") or ""),
            "Successore": str(succ_id),
            "Successore Unique ID": str(getattr(succ, "uid", "") or ""),
            "Relazione": relazione,
            "Lag giorni": lag_giorni,
            "Testo predecessore": testo,
        })
    return predecessori, dettaglio


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
        predecessori, dettaglio_dipendenze = costruisci_predecessori(project)
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
                COL_UNIQUE_ID: str(getattr(task, "uid", "") or ""),
                COL_PREDECESSORI: ";".join(predecessori.get(task_id, [])),
                COL_NOME: str(task.name or "").strip(),
                COL_DURATA: str(task.duration or ""),
                COL_INIZIO: pd.to_datetime(task.start, errors="coerce"),
                COL_FINE: pd.to_datetime(task.finish, errors="coerce"),
            })

        df = pd.DataFrame(righe, columns=[
            COL_MACRO, COL_ID, COL_UNIQUE_ID, COL_PREDECESSORI, COL_NOME,
            COL_DURATA, COL_INIZIO, COL_FINE,
        ])
        df[COL_MACRO] = df[COL_MACRO].fillna("").astype(str).str.strip()
        df.loc[df[COL_MACRO] == "", COL_MACRO] = "Unclassified"
        df.attrs["summary_escluse"] = numero_summary_escluse
        df.attrs["lettore_mpp"] = READER_NAME
        df.attrs["numero_dipendenze"] = len(dettaglio_dipendenze)
        df.attrs["dettaglio_dipendenze"] = dettaglio_dipendenze
        return df
    except Exception as exc:
        raise RuntimeError(f"Unable to read MPP file '{nome_file}': {exc}") from exc
    finally:
        if percorso_temporaneo is not None:
            try:
                percorso_temporaneo.unlink(missing_ok=True)
            except Exception:
                pass


