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

READER_NAME = "MPXJ"
REQUIREMENTS = "mpxj, JPype1 and Java 11+"

def _java_value(value, default=None):
    """Converte in modo tollerante un valore Java/JPype in un valore Python."""
    if value is None:
        return default
    text = str(value).strip()
    return default if text in {"", "None", "null"} else text


def _java_datetime(value):
    """Converte LocalDate/LocalDateTime MPXJ in Timestamp pandas."""
    if value is None:
        return pd.NaT
    try:
        return pd.Timestamp(
            int(value.getYear()), int(value.getMonthValue()), int(value.getDayOfMonth()),
            int(value.getHour()) if hasattr(value, "getHour") else 0,
            int(value.getMinute()) if hasattr(value, "getMinute") else 0,
            int(value.getSecond()) if hasattr(value, "getSecond") else 0,
        )
    except Exception:
        return pd.to_datetime(str(value), errors="coerce")


def _codice_relazione_mpxj(relation_type) -> str:
    nome = str(relation_type or "FINISH_START").upper().replace("-", "_")
    mapping = {
        "FINISH_START": "FI", "FINISH_TO_START": "FI", "FS": "FI",
        "START_START": "II", "START_TO_START": "II", "SS": "II",
        "FINISH_FINISH": "FF", "FINISH_TO_FINISH": "FF",
        "START_FINISH": "IF", "START_TO_FINISH": "IF", "SF": "IF",
    }
    return mapping.get(nome, "FI")


def _lag_giorni_mpxj(relation, project) -> float:
    """Converte il lag MPXJ in giorni usando i parametri calendario del progetto."""
    try:
        lag = relation.getLag()
        if lag is None:
            return 0.0
        try:
            from org.mpxj import TimeUnit
            convertito = lag.convertUnits(TimeUnit.DAYS, project.getProjectProperties())
            return float(convertito.getDuration())
        except Exception:
            valore = float(lag.getDuration())
            unita = str(lag.getUnits()).upper()
            if "MINUTE" in unita:
                return valore / 480.0
            if "HOUR" in unita:
                return valore / 8.0
            if "WEEK" in unita:
                return valore * 5.0
            if "MONTH" in unita:
                return valore * 20.0
            return valore
    except Exception:
        return 0.0



def _prisma_macro_mpxj(project, task) -> str:
    """Cerca il campo personalizzato Prisma Macro Activity; fallback Text10."""
    campi = []
    try:
        campi = list(project.getCustomFields())
    except Exception:
        pass
    fallback = ""
    for field in campi:
        try:
            alias = str(field.getAlias() or "")
            field_type = field.getFieldType()
            field_name = str(field_type or "")
            chiave = f"{alias} {field_name}".lower().replace(" ", "").replace("_", "")
            valore = task.getCachedValue(field_type)
            valore = "" if valore is None else str(valore).strip()
            if "prismamacroactivity" in chiave and valore:
                return valore
            if "text10" in chiave and valore and not fallback:
                fallback = valore
        except Exception:
            continue
    return fallback or "Unclassified"


def _predecessori_mpxj(project) -> tuple[dict[int, list[str]], list[dict]]:
    """
    Estrae le dipendenze usando i metodi espliciti MPXJ:
    getPredecessorTask() e getSuccessorTask().

    Non usa più getSourceTask()/getTargetTask(), perché in MPXJ recenti sono
    deprecati e possono portare a interpretare al contrario gli estremi.
    """
    predecessori = defaultdict(list)
    dettaglio = []
    archi_visti = set()

    for task in project.getTasks():
        if task is None:
            continue
        try:
            relazioni = task.getPredecessors()
        except Exception:
            relazioni = []

        for relation in relazioni or []:
            try:
                pred = relation.getPredecessorTask()
                succ = relation.getSuccessorTask()
            except Exception:
                # Compatibilità con eventuali versioni MPXJ meno recenti.
                try:
                    pred = relation.getSourceTask()
                    succ = relation.getTargetTask()
                except Exception:
                    continue

            if pred is None or succ is None:
                continue
            if pred.getID() is None or succ.getID() is None:
                continue

            pred_id = int(str(pred.getID()))
            succ_id = int(str(succ.getID()))
            if pred_id == succ_id:
                continue

            relazione = _codice_relazione_mpxj(relation.getType())
            lag_giorni = _lag_giorni_mpxj(relation, project)
            chiave = (pred_id, succ_id, relazione, round(lag_giorni, 8))
            if chiave in archi_visti:
                continue
            archi_visti.add(chiave)

            lag_testo = formatta_lag(lag_giorni)
            testo = (
                str(pred_id)
                if relazione == "FI" and not lag_testo
                else f"{pred_id}{relazione}{lag_testo}"
            )
            predecessori[succ_id].append(testo)
            dettaglio.append({
                "Predecessore": str(pred_id),
                "Predecessore Unique ID": _java_value(pred.getUniqueID(), ""),
                "Successore": str(succ_id),
                "Successore Unique ID": _java_value(succ.getUniqueID(), ""),
                "Relazione": relazione,
                "Lag giorni": lag_giorni,
                "Testo predecessore": testo,
            })

    return predecessori, dettaglio


def carica_schedule_mpp(contenuto: bytes, nome_file: str) -> pd.DataFrame:
    """Legge MPP con MPXJ/JPype, senza Aspose e senza dipendenza da libssl1.1."""
    try:
        import jpype
        import mpxj  # noqa: F401: registra il classpath MPXJ
    except ImportError as exc:
        raise RuntimeError(
            "Per leggere i file MPP installa: python -m pip install mpxj JPype1"
        ) from exc

    percorso_temporaneo = None
    try:
        if not jpype.isJVMStarted():
            try:
                jpype.startJVM(convertStrings=True)
            except Exception as exc:
                raise RuntimeError(
                    "Impossibile avviare Java. Installa un JDK/JRE 11 o successivo e "
                    "verifica JAVA_HOME. Dettaglio: " + str(exc)
                ) from exc

        from org.mpxj.reader import UniversalProjectReader

        with tempfile.NamedTemporaryFile(suffix=".mpp", delete=False) as tmp:
            tmp.write(contenuto)
            percorso_temporaneo = Path(tmp.name)

        project = UniversalProjectReader().read(str(percorso_temporaneo))
        predecessori, dettaglio_dipendenze = _predecessori_mpxj(project)
        righe = []
        numero_summary_escluse = 0

        for task in project.getTasks():
            if task is None or task.getID() is None:
                continue
            task_id = int(str(task.getID()))
            if task_id == 0:
                continue
            try:
                is_summary = bool(task.getSummary())
            except Exception:
                try:
                    is_summary = bool(task.getChildTasks())
                except Exception:
                    is_summary = False
            if is_summary:
                numero_summary_escluse += 1
                continue

            righe.append({
                COL_MACRO: _prisma_macro_mpxj(project, task),
                COL_ID: str(task_id),
                COL_UNIQUE_ID: _java_value(task.getUniqueID(), ""),
                COL_PREDECESSORI: ";".join(predecessori.get(task_id, [])),
                COL_NOME: _java_value(task.getName(), ""),
                COL_DURATA: _java_value(task.getDuration(), ""),
                COL_INIZIO: _java_datetime(task.getStart()),
                COL_FINE: _java_datetime(task.getFinish()),
            })

        df = pd.DataFrame(righe, columns=[
            COL_MACRO, COL_ID, COL_UNIQUE_ID, COL_PREDECESSORI, COL_NOME,
            COL_DURATA, COL_INIZIO, COL_FINE,
        ])
        df[COL_MACRO] = df[COL_MACRO].fillna("Unclassified").astype(str).str.strip()
        df.loc[df[COL_MACRO] == "", COL_MACRO] = "Unclassified"
        df.attrs["summary_escluse"] = numero_summary_escluse
        df.attrs["lettore_mpp"] = READER_NAME
        df.attrs["numero_dipendenze"] = len(dettaglio_dipendenze)
        df.attrs["dettaglio_dipendenze"] = dettaglio_dipendenze
        return df
    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Unable to read MPP file '{nome_file}' with MPXJ: {exc}") from exc
    finally:
        if percorso_temporaneo is not None:
            try:
                percorso_temporaneo.unlink(missing_ok=True)
            except Exception:
                pass


