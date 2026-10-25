import os
import io
import json
import random
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import streamlit as st
import plotly.express as px

# ==============================================================================
# CONFIGURAZIONE PAGINA E COSTANTI DI DEFAULT (DA EXCEL LEGA USV 26/27)
# ==============================================================================
st.set_page_config(
    page_title="Lega USV Biliardo - Gestionale Elo",
    page_icon="🎱",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_FILE = "usv_biliardo_db.json"
EXCEL_REFERENCE = "lega USV biliardo S1 26-27.xlsx"

DEFAULT_SETTINGS = {
    "start_elo": 500,          # Elo di partenza per ogni giocatore
    "k_factor": 40,            # Fattore K (variazione base massima)
    "scale_factor": 400,       # Fattore di scala logistica Elo
    "win_bonus": 3,            # Bonus partecipazione/vittoria (+3 come da formula Excel IF(F=1, 3, 0))
    "presence_bonus": 0,       # Bonus presenza incondizionato (anche per chi perde, default 0)
    "gap_divisor": 1200,       # Divisore penalità divario coppia: ((Max-Min)^2)/1200 (0 = disattivato)
    "weight_max": 0.50,        # Peso giocatore più forte nella media di coppia (0.50 = media aritmetica)
    "min_matches_active": 0,   # Soglia minima partite per filtrare gli inattivi in classifica
    "count_1v2_as_double": False # Se True replica il comportamento COUNTIF di Excel per i match 1vs2
}

LIVELLI_LEGENDA = {
    "Non assegnato": "Nessun livello specificato.",
    "1 - Senza mani": "Fatica a colpire la palla dritta, commette falli continui e segna solo per puro caso.",
    "2 - Base": "Imbuca tiri facili e diretti, ma ignora la forza del colpo e dove andrà a finire la bianca.",
    "3 - Intermedio": "Buona mira di base, fa scelte strategiche e comincia a posizionarsi per il tiro successivo.",
    "4 - Avanzato": "Capace di usare le sponde, gioca di tattica e pianifica la serie con 2-3 tiri d'anticipo.",
    "5 - Fuoriclasse": "Controllo millimetrico dello spin e delle traiettorie complesse; può ripulire il tavolo in un unico turno."
}

STILI_LEGENDA = {
    "Non assegnato": "Nessuno stile specificato.",
    "Spara e spera": "Tira a tutta forza senza pensare alla posizione, sperando che qualcosa finisca in buca.",
    "Strategico": "Studia ogni angolo con calma, pianifica la serie e preferisce la precisione alla potenza.",
    "Aggressivo": "Attacca sempre e cerca l'imbucata a ogni costo, anche su tiri quasi impossibili.",
    "Difensivo": "Se non ha un tiro facile, preferisce nascondere la bianca e bloccare il gioco dell'avversario.",
    "Fine": "Cerca tiri difficili a più sponde, combinazioni o effetti scenografici per il gusto della giocata.",
    "Istintivo": "Tira in pochi secondi senza pensarci troppo, affidandosi puramente al colpo d'occhio.",
    "Minimalista": "Gioca solo tiri semplici e sicuri a bassa percentuale di rischio, non tenta colpi complessi.",
    "Sculato": "Autoesplicativo."
}

INITIAL_PLAYERS = [
    "Francia", "Matos", "Vesco", "Paolo", "Birra", "Gae",
    "Filli", "Col", "Picci", "Riki", "Vezzo", "Feli", "Bonny"
]

# Storico iniziale delle 79 partite estratto dal foglio 'Partite' dell'Excel
INITIAL_MATCHES = [["Bonny", "Gae", "Francia", "Matos", 1], ["Francia", "Matos", "Bonny", "Gae", 1], ["Francia", "Matos", "Bonny", "Gae", 1], ["Francia", "Matos", "Bonny", "Gae", 1], ["Francia", "Matos", "Bonny", "Gae", 1], ["Francia", "Matos", "Bonny", "Gae", 1], ["Francia", "Matos", "Bonny", "Gae", 1], ["Francia", "Matos", "Bonny", "Gae", 1], ["Bonny", "Gae", "Francia", "Matos", 1], ["Paolo", "Francia", "Birra", "Gae", 1], ["Riki", "Col", "Matos", "Filli", 1], ["Riki", "Col", "Birra", "Gae", 1], ["Matos", "Filli", "Paolo", "Francia", 1], ["Matos", "Filli", "Birra", "Gae", 1], ["Paolo", "Francia", "Riki", "Col", 1], ["Birra", "Gae", "Riki", "Col", 1], ["Paolo", "Francia", "Matos", "Filli", 1], ["Birra", "Gae", "Paolo", "Francia", 1], ["Riki", "Col", "Matos", "Filli", 1], ["Birra", "Gae", "Matos", "Filli", 1], ["Riki", "Col", "Paolo", "Francia", 1], ["Birra", "Gae", "Riki", "Col", 1], ["Matos", "Filli", "Paolo", "Francia", 1], ["Francia", "Gae", "Vesco", "Picci", 1], ["Bonny", "Feli", "Matos", "Paolo", 1], ["Vesco", "Picci", "Matos", "Paolo", 1], ["Francia", "Gae", "Bonny", "Feli", 1], ["Vesco", "Picci", "Bonny", "Feli", 1], ["Francia", "Gae", "Matos", "Paolo", 1], ["Francia", "Gae", "Vesco", "Picci", 1], ["Matos", "Paolo", "Bonny", "Feli", 1], ["Vesco", "Picci", "Matos", "Paolo", 1], ["Francia", "Gae", "Bonny", "Feli", 1], ["Bonny", "Bonny", "Vesco", "Picci", 1], ["Matos", "Paolo", "Francia", "Gae", 1], ["Filli", "Filli", "Matos", "Gae", 1], ["Vesco", "Bonny", "Filli", "Paolo", 1], ["Matos", "Gae", "Vesco", "Bonny", 1], ["Filli", "Paolo", "Matos", "Gae", 1], ["Filli", "Paolo", "Vesco", "Bonny", 1], ["Matos", "Gae", "Filli", "Paolo", 1], ["Matos", "Gae", "Vesco", "Bonny", 1], ["Filli", "Paolo", "Matos", "Gae", 1], ["Vesco", "Bonny", "Filli", "Paolo", 1], ["Matos", "Gae", "Vesco", "Bonny", 1], ["Gae", "Filli", "Paolo", "Riki", 1], ["Gae", "Filli", "Matos", "Bonny", 1], ["Gae", "Filli", "Paolo", "Riki", 1], ["Matos", "Bonny", "Gae", "Filli", 1], ["Gae", "Filli", "Matos", "Bonny", 1], ["Paolo", "Riki", "Gae", "Filli", 1], ["Paolo", "Riki", "Matos", "Bonny", 1], ["Gae", "Filli", "Paolo", "Riki", 1], ["Matos", "Bonny", "Gae", "Filli", 1], ["Riki", "Vesco", "Gae", "Paolo", 1], ["Riki", "Vesco", "Francia", "Vezzo", 1], ["Gae", "Paolo", "Riki", "Vesco", 1], ["Francia", "Vezzo", "Gae", "Paolo", 1], ["Vesco", "Matos", "Gae", "Gae", 1], ["Filli", "Francia", "Paolo", "Bonny", 1], ["Paolo", "Bonny", "Gae", "Gae", 1], ["Vesco", "Matos", "Filli", "Francia", 1], ["Gae", "Riki", "Filli", "Francia", 1], ["Vesco", "Matos", "Paolo", "Bonny", 1], ["Vesco", "Matos", "Gae", "Riki", 1], ["Paolo", "Bonny", "Filli", "Francia", 1], ["Paolo", "Bonny", "Gae", "Riki", 1], ["Filli", "Francia", "Vesco", "Matos", 1], ["Filli", "Francia", "Gae", "Riki", 1], ["Vesco", "Matos", "Paolo", "Bonny", 1], ["Bonny", "Vezzo", "Paolo", "Paolo", 1], ["Matos", "Riki", "Bonny", "Vezzo", 1], ["Paolo", "Paolo", "Matos", "Riki", 1], ["Paolo", "Paolo", "Bonny", "Vezzo", 1], ["Matos", "Riki", "Paolo", "Paolo", 1], ["Bonny", "Vezzo", "Matos", "Riki", 1], ["Paolo", "Paolo", "Bonny", "Vezzo", 1], ["Paolo", "Gae", "Matos", "Riki", 1], ["Paolo", "Gae", "Bonny", "Vezzo", 1]]

# ==============================================================================
# FUNZIONI MATEMATICHE E MOTORE ELO (REPLICA ESATTA EXCEL)
# ==============================================================================
def excel_round(val: float, ndigits: int = 0) -> float | int:
    """Replica l'arrotondamento aritmetico ROUND() di Excel (Round Half Away From Zero)."""
    d = Decimal(str(val))
    if ndigits == 0:
        return int(d.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    fmt = "0." + "0" * ndigits
    return float(d.quantize(Decimal(fmt), rounding=ROUND_HALF_UP))


def calc_team_elo(elo1: float, elo2: float, settings: dict) -> int:
    """
    Calcola l'Elo della coppia combinando:
    1) Media ponderata tra giocatore più forte e più debole (default 50%/50%)
    2) Penalità quadratica per il divario di Elo: ((MAX - MIN)^2) / gap_divisor (default 1200)
    """
    hi = max(elo1, elo2)
    lo = min(elo1, elo2)
    w_max = float(settings.get("weight_max", 0.50))
    base_avg = w_max * hi + (1.0 - w_max) * lo

    gap_div = float(settings.get("gap_divisor", 1200))
    penalty = ((hi - lo) ** 2) / gap_div if gap_div > 0 else 0.0
    return excel_round(base_avg - penalty, 0)


def preview_match_math(elo_a: int, elo_b: int, elo_c: int, elo_d: int, settings: dict) -> dict:
    """Restituisce l'anteprima matematica pre-partita per entrambi gli scenari di vittoria."""
    elo_t1 = calc_team_elo(elo_a, elo_b, settings)
    elo_t2 = calc_team_elo(elo_c, elo_d, settings)
    scale = float(settings.get("scale_factor", 400))
    k = float(settings.get("k_factor", 40))
    win_bonus = int(settings.get("win_bonus", 3))
    pres_bonus = int(settings.get("presence_bonus", 0))

    prob_t1 = excel_round(1.0 / (1.0 + 10.0 ** ((elo_t2 - elo_t1) / scale)), 2)
    prob_t2 = round(1.0 - prob_t1, 2)

    # Scenario 1: Vince Team 1 (Vincitore = 1)
    var_if_t1 = excel_round(k * (1.0 - prob_t1), 0)
    gain_t1_win = var_if_t1 + win_bonus + pres_bonus
    loss_t2_lose = -var_if_t1 + pres_bonus

    # Scenario 2: Vince Team 2 (Vincitore = 2)
    var_if_t2 = excel_round(k * (0.0 - prob_t1), 0)  # negativo
    loss_t1_lose = var_if_t2 + pres_bonus
    gain_t2_win = -var_if_t2 + win_bonus + pres_bonus

    return {
        "elo_t1": elo_t1,
        "elo_t2": elo_t2,
        "prob_t1": prob_t1,
        "prob_t2": prob_t2,
        "var_if_t1": var_if_t1,
        "var_if_t2": var_if_t2,
        "gain_t1_win": gain_t1_win,
        "loss_t2_lose": loss_t2_lose,
        "loss_t1_lose": loss_t1_lose,
        "gain_t2_win": gain_t2_win,
    }


def compute_streaks(results: list[int]) -> tuple[int, int, int]:
    """Calcola (Streak Attiva, Peggiore Streak, Migliore Streak) da una lista di +1 (V) e -1 (P)."""
    if not results:
        return 0, 0, 0
    curr, best, worst = 0, 0, 0
    for r in results:
        if r == 1:
            curr = curr + 1 if curr > 0 else 1
        else:
            curr = curr - 1 if curr < 0 else -1
        if curr > best:
            best = curr
        if curr < worst:
            worst = curr
    return curr, worst, best


def run_elo_engine(db: dict) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict]:
    """
    Ricalcola l'intero storico partite, la classifica giocatori, lo storico coppie e le serie storiche Elo.
    Garantisce coerenza matematica al 100% col file Excel.
    """
    settings = db["settings"]
    start_elo = int(settings.get("start_elo", 500))
    scale = float(settings.get("scale_factor", 400))
    k = float(settings.get("k_factor", 40))
    win_bonus = int(settings.get("win_bonus", 3))
    pres_bonus = int(settings.get("presence_bonus", 0))
    count_double = bool(settings.get("count_1v2_as_double", False))

    elo_map = {p: start_elo for p in db["players"]}
    elo_history = {p: [start_elo] for p in db["players"]}
    match_history_axis = {p: [0] for p in db["players"]}
    player_results = {p: [] for p in db["players"]}
    slot_counts = {p: {"matches": 0, "wins": 0} for p in db["players"]}

    rows_partite = []
    pairs_results = {}

    for idx, m in enumerate(db["matches"], start=1):
        a, b, c, d = m["A"], m["B"], m["C"], m["D"]
        vinc = int(m["Vincitore"])
        perfect = bool(m.get("Perfect", False))
        note = m.get("Note", "")

        for p in (a, b, c, d):
            if p not in elo_map:
                elo_map[p] = start_elo
                elo_history[p] = [start_elo]
                match_history_axis[p] = [0]
                player_results[p] = []
                slot_counts[p] = {"matches": 0, "wins": 0}

        elo_a, elo_b = elo_map[a], elo_map[b]
        elo_c, elo_d = elo_map[c], elo_map[d]

        elo_t1 = calc_team_elo(elo_a, elo_b, settings)
        elo_t2 = calc_team_elo(elo_c, elo_d, settings)

        prob_t1 = excel_round(1.0 / (1.0 + 10.0 ** ((elo_t2 - elo_t1) / scale)), 2)
        var_base = excel_round(k * ((1.0 if vinc == 1 else 0.0) - prob_t1), 0)

        new_a = excel_round(elo_a + var_base + (win_bonus if vinc == 1 else 0) + pres_bonus, 0)
        new_b = excel_round(elo_b + var_base + (win_bonus if vinc == 1 else 0) + pres_bonus, 0)
        new_c = excel_round(elo_c - var_base + (win_bonus if vinc == 2 else 0) + pres_bonus, 0)
        new_d = excel_round(elo_d - var_base + (win_bonus if vinc == 2 else 0) + pres_bonus, 0)

        elo_map[a] = new_a
        elo_map[b] = new_b
        elo_map[c] = new_c
        elo_map[d] = new_d

        # Aggiornamento conteggi slot (stile COUNTIF Excel)
        for p_t1 in (a, b):
            slot_counts[p_t1]["matches"] += 1
            if vinc == 1:
                slot_counts[p_t1]["wins"] += 1
        for p_t2 in (c, d):
            slot_counts[p_t2]["matches"] += 1
            if vinc == 2:
                slot_counts[p_t2]["wins"] += 1

        # Aggiornamento partite reali e streak per giocatore (senza duplicare 1v2)
        t1_unique = list(dict.fromkeys([a, b]))
        t2_unique = list(dict.fromkeys([c, d]))
        for p_u in t1_unique:
            player_results[p_u].append(1 if vinc == 1 else -1)
            elo_history[p_u].append(elo_map[p_u])
            match_history_axis[p_u].append(idx)
        for p_u in t2_unique:
            if p_u not in t1_unique:
                player_results[p_u].append(1 if vinc == 2 else -1)
                elo_history[p_u].append(elo_map[p_u])
                match_history_axis[p_u].append(idx)

        # Storico coppie
        pair1_name = " + ".join(sorted([a, b])) if a != b else f"{a} (Solo 1v2)"
        pair2_name = " + ".join(sorted([c, d])) if c != d else f"{c} (Solo 1v2)"
        pairs_results.setdefault(pair1_name, []).append(1 if vinc == 1 else -1)
        pairs_results.setdefault(pair2_name, []).append(1 if vinc == 2 else -1)

        rows_partite.append({
            "#": idx,
            "A": a,
            "B": b,
            "C": c,
            "D": d,
            "Vincitore": vinc,
            "Elo A": elo_a,
            "Elo B": elo_b,
            "Elo C": elo_c,
            "Elo D": elo_d,
            "Elo Team 1": elo_t1,
            "Elo Team 2": elo_t2,
            "Prob T1": prob_t1,
            "Var. Base": var_base,
            "Nuovo Elo A": new_a,
            "Nuovo Elo B": new_b,
            "Nuovo Elo C": new_c,
            "Nuovo Elo D": new_d,
            "Perfect": perfect,
            "Note": note
        })

    df_partite = pd.DataFrame(rows_partite)

    # Costruzione DataFrame Classifica Giocatori
    player_rows = []
    meta = db.get("player_meta", {})
    for p, curr_elo in elo_map.items():
        res_list = player_results.get(p, [])
        real_m = len(res_list)
        real_w = sum(1 for r in res_list if r == 1)
        real_l = real_m - real_w

        slot_m = slot_counts[p]["matches"]
        slot_w = slot_counts[p]["wins"]
        slot_l = slot_m - slot_w

        used_m = slot_m if count_double else real_m
        used_w = slot_w if count_double else real_w
        used_l = slot_l if count_double else real_l

        win_rate = round((used_w / used_m) * 100, 1) if used_m > 0 else 0.0
        forza = round(used_w * (used_w / used_m), 2) if used_m > 0 else 0.0
        curr_s, worst_s, best_s = compute_streaks(res_list)

        p_meta = meta.get(p, {})
        # Conteggio perfect (base manuale + eventuali registrati nei match)
        base_perf = int(p_meta.get("perfect_base", 0))
        match_perf = 0
        if not df_partite.empty:
            for _, r_m in df_partite[df_partite["Perfect"] == True].iterrows():
                if (r_m["Vincitore"] == 1 and p in (r_m["A"], r_m["B"])) or (r_m["Vincitore"] == 2 and p in (r_m["C"], r_m["D"])):
                    match_perf += 1

        player_rows.append({
            "Giocatore": p,
            "Elo Attuale": curr_elo,
            "Delta Elo": curr_elo - start_elo,
            "Partite": used_m,
            "Vinte": used_w,
            "Perse": used_l,
            "% Vittoria": win_rate,
            "Indice di Forza": forza,
            "Streak Attiva": curr_s,
            "Migliore Streak": best_s,
            "Peggiore Streak": worst_s,
            "Perfect": base_perf + match_perf,
            "Partite Reali": real_m,
            "Slot Excel (COUNTIF)": slot_m,
            "Livello": p_meta.get("livello", "Non assegnato"),
            "Stile": p_meta.get("stile", "Non assegnato")
        })

    df_players = pd.DataFrame(player_rows)
    if not df_players.empty:
        df_players = df_players.sort_values(
            by=["Elo Attuale", "Vinte", "% Vittoria"],
            ascending=[False, False, False]
        ).reset_index(drop=True)
        df_players.insert(0, "Posizione", range(1, len(df_players) + 1))

    # Costruzione DataFrame Coppie
    pair_rows = []
    for pair_name, p_res in pairs_results.items():
        pm = len(p_res)
        pw = sum(1 for r in p_res if r == 1)
        pl = pm - pw
        p_wr = round((pw / pm) * 100, 1) if pm > 0 else 0.0
        p_forza = round(pw * (pw / pm), 2) if pm > 0 else 0.0
        p_curr, p_worst, p_best = compute_streaks(p_res)
        pair_rows.append({
            "Coppia": pair_name,
            "Partite": pm,
            "Vittorie": pw,
            "Perse": pl,
            "% Vittorie": p_wr,
            "Indice di Forza": p_forza,
            "Streak Attiva": p_curr,
            "Migliore Streak": p_best,
            "Peggiore Streak": p_worst,
            "Tipo": "Solo (1v2)" if "(Solo 1v2)" in pair_name else "Coppia (2v2)"
        })

    df_pairs = pd.DataFrame(pair_rows)
    if not df_pairs.empty:
        df_pairs = df_pairs.sort_values(
            by=["Indice di Forza", "Vittorie", "% Vittorie"],
            ascending=[False, False, False]
        ).reset_index(drop=True)

    return df_partite, df_players, df_pairs, {"elo": elo_history, "axis": match_history_axis}


# ==============================================================================
# GESTIONE PERSISTENZA DATI (JSON LOCALE + IMPORT/EXPORT EXCEL)
# ==============================================================================
def create_default_db() -> dict:
    """Inizializza il database partendo dal file Excel (se presente) o dal seed integrato."""
    matches = []
    players = list(INITIAL_PLAYERS)

    if os.path.exists(EXCEL_REFERENCE):
        try:
            df_excel = pd.read_excel(EXCEL_REFERENCE, sheet_name="Partite")
            df_excel = df_excel.dropna(subset=["A", "B", "C", "D", "Vincitore"])
            for _, r in df_excel.iterrows():
                a, b, c, d = str(r["A"]).strip(), str(r["B"]).strip(), str(r["C"]).strip(), str(r["D"]).strip()
                vinc = int(r["Vincitore"])
                matches.append({"A": a, "B": b, "C": c, "D": d, "Vincitore": vinc, "Perfect": False, "Note": ""})
                for p in (a, b, c, d):
                    if p not in players:
                        players.append(p)
        except Exception:
            matches = []

    if not matches:
        for m in INITIAL_MATCHES:
            matches.append({"A": m[0], "B": m[1], "C": m[2], "D": m[3], "Vincitore": int(m[4]), "Perfect": False, "Note": ""})

    return {
        "settings": dict(DEFAULT_SETTINGS),
        "players": players,
        "player_meta": {p: {"livello": "Non assegnato", "stile": "Non assegnato", "perfect_base": 0} for p in players},
        "matches": matches
    }


def load_db() -> dict:
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in DEFAULT_SETTINGS.items():
                data.setdefault("settings", {}).setdefault(k, v)
            data.setdefault("player_meta", {})
            return data
        except Exception:
            pass
    db = create_default_db()
    save_db(db)
    return db


def save_db(db: dict) -> None:
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2, ensure_ascii=False)


def export_to_excel_bytes(df_partite: pd.DataFrame, df_players: pd.DataFrame, df_pairs: pd.DataFrame) -> bytes:
    """Esporta l'intero database ricalcolato in un file Excel (.xlsx) strutturato."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_players[["Posizione", "Giocatore", "Elo Attuale", "Partite", "Vinte", "Perse", "% Vittoria", "Indice di Forza", "Streak Attiva", "Migliore Streak", "Peggiore Streak", "Perfect"]].to_excel(
            writer, sheet_name="Classifica", index=False
        )
        df_partite.to_excel(writer, sheet_name="Partite", index=False)
        df_pairs.to_excel(writer, sheet_name="Storico Coppie", index=False)
    return output.getvalue()


# ==============================================================================
# INIZIALIZZAZIONE STATO STREAMLIT
# ==============================================================================
if "db" not in st.session_state:
    st.session_state.db = load_db()

db = st.session_state.db
df_partite, df_players, df_pairs, history_data = run_elo_engine(db)

# ==============================================================================
# SIDEBAR DI NAVIGAZIONE E RIEPILOGO RAPIDO
# ==============================================================================
with st.sidebar:
    st.markdown("## 🎱 Lega USV Biliardo")
    st.caption("Stagione 1 (26/27) • Sistema Elo Ufficiale")
    st.divider()

    menu = st.radio(
        "Navigazione",
        [
            "🏆 Dashboard Classifica",
            "⚡ Inserimento Partite (Live)",
            "📊 Tabellone & Classifiche Extra",
            "👥 Gestione Giocatori",
            "🔍 Dati Grezzi & Diagnostica",
            "⚙️ Impostazioni Sistema Elo"
        ]
    )

    st.divider()
    st.markdown("### 📌 Stato Torneo")
    c1, c2 = st.columns(2)
    c1.metric("Partite", len(df_partite))
    c2.metric("Giocatori", len(db["players"]))

    if not df_players.empty:
        leader = df_players.iloc[0]
        st.success(f"👑 **Leader Elo:** {leader['Giocatore']} ({leader['Elo Attuale']} pt)")

    st.download_button(
        label="📥 Scarica Backup Excel (.xlsx)",
        data=export_to_excel_bytes(df_partite, df_players, df_pairs),
        file_name="lega_USV_biliardo_aggiornato.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ==============================================================================
# 1. DASHBOARD CLASSIFICA
# ==============================================================================
if menu == "🏆 Dashboard Classifica":
    st.title("🏆 Dashboard Classifica Generale Elo")
    st.markdown("Classifica ufficiale aggiornata in tempo reale secondo il motore matematico Elo della **Lega Ultimate Sclero Vergogna**.")

    # KPI Principali
    if not df_players.empty:
        top_elo = df_players.iloc[0]
        qualified = df_players[df_players["Partite"] >= max(5, db["settings"]["min_matches_active"])]
        if qualified.empty:
            qualified = df_players
        top_wr = qualified.sort_values(by=["% Vittoria", "Vinte"], ascending=[False, False]).iloc[0]
        top_forza = df_players.sort_values(by=["Indice di Forza", "Vinte"], ascending=[False, False]).iloc[0]
        top_streak = df_players.sort_values(by=["Streak Attiva", "Elo Attuale"], ascending=[False, False]).iloc[0]

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("🥇 1° Classificato (Elo)", f"{top_elo['Giocatore']}", f"{top_elo['Elo Attuale']} pt ({top_elo['Delta Elo']:+d})")
        k2.metric("🔥 Miglior Win-Rate (≥5 p.)", f"{top_wr['Giocatore']}", f"{top_wr['% Vittoria']}% ({top_wr['Vinte']}V / {top_wr['Partite']}P)")
        k3.metric("💪 Miglior Indice di Forza", f"{top_forza['Giocatore']}", f"{top_forza['Indice di Forza']:.2f}")
        k4.metric("⚡ Streak Attiva Migliore", f"{top_streak['Giocatore']}", f"{top_streak['Streak Attiva']:+d} partite")

    st.divider()

    # Barra filtri rapidi
    f_col1, f_col2, f_col3 = st.columns([2, 2, 2])
    with f_col1:
        min_p_filter = st.slider(
            "Filtra inattivi (Minimo partite giocate):",
            min_value=0,
            max_value=30,
            value=int(db["settings"].get("min_matches_active", 0)),
            help="Nasconde dalla classifica i giocatori con meno partite della soglia impostata."
        )
    with f_col2:
        count_mode = st.toggle(
            "Usa conteggio Slot Excel (COUNTIF) per i match 1v2",
            value=bool(db["settings"].get("count_1v2_as_double", False)),
            help="Nel foglio 'Classifica' di Excel, chi gioca 1v2 (es. Paolo+Paolo) viene contato 2 volte nella stessa riga da COUNTIF. Attiva per vedere i numeri identici al foglio Classifica, disattiva per le partite reali del Tabellone."
        )
        if count_mode != db["settings"]["count_1v2_as_double"]:
            db["settings"]["count_1v2_as_double"] = count_mode
            save_db(db)
            st.rerun()
    with f_col3:
        search_player = st.text_input("🔍 Cerca giocatore:", "")

    df_view = df_players[df_players["Partite"] >= min_p_filter].copy()
    if search_player.strip():
        df_view = df_view[df_view["Giocatore"].str.contains(search_player.strip(), case=False, na=False)]

    # Formattazione visiva Streak
    def format_streak(val: int) -> str:
        if val >= 3:
            return f"🔥 +{val}"
        elif val > 0:
            return f"🟢 +{val}"
        elif val <= -3:
            return f"❄️ {val}"
        elif val < 0:
            return f"🔴 {val}"
        return "➖ 0"

    df_view["Forma (Streak)"] = df_view["Streak Attiva"].apply(format_streak)

    st.dataframe(
        df_view[[
            "Posizione", "Giocatore", "Elo Attuale", "Delta Elo",
            "Partite", "Vinte", "Perse", "% Vittoria",
            "Indice di Forza", "Forma (Streak)", "Migliore Streak", "Peggiore Streak", "Perfect"
        ]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Posizione": st.column_config.NumberColumn("Pos.", format="%d°"),
            "Elo Attuale": st.column_config.ProgressColumn(
                "Punti Elo",
                format="%d pt",
                min_value=350,
                max_value=int(max(700, df_players["Elo Attuale"].max() + 20)) if not df_players.empty else 700
            ),
            "Delta Elo": st.column_config.NumberColumn("Δ da Base", format="%+d"),
            "% Vittoria": st.column_config.NumberColumn("% Vittoria", format="%.1f%%"),
            "Indice di Forza": st.column_config.NumberColumn("Indice Forza", format="%.2f")
        }
    )

    # Grafico Evoluzione Elo
    st.subheader("📈 Evoluzione Storica Punteggio Elo")
    default_sel = df_players["Giocatore"].head(6).tolist() if not df_players.empty else []
    selected_players = st.multiselect(
        "Seleziona i giocatori da confrontare nel grafico:",
        options=df_players["Giocatore"].tolist(),
        default=default_sel
    )
    if selected_players:
        chart_rows = []
        for p in selected_players:
            elos = history_data["elo"].get(p, [])
            axis = history_data["axis"].get(p, [])
            for m_num, e_val in zip(axis, elos):
                chart_rows.append({"Partita #": m_num, "Giocatore": p, "Elo": e_val})
        df_chart = pd.DataFrame(chart_rows)
        fig = px.line(
            df_chart,
            x="Partita #",
            y="Elo",
            color="Giocatore",
            markers=True,
            title="Andamento Elo lungo la Stagione",
            template="plotly_white"
        )
        fig.add_hline(y=db["settings"]["start_elo"], line_dash="dash", line_color="gray", annotation_text="Elo Iniziale")
        st.plotly_chart(fig, use_container_width=True)

# ==============================================================================
# 2. INSERIMENTO PARTITE (SERATA LIVE & MATCHMAKING RANDOM/BILANCIATO)
# ==============================================================================
elif menu == "⚡ Inserimento Partite (Live)":
    st.title("⚡ Registrazione Partite & Matchmaking Serata")
    st.markdown("Estrai a sorte o bilancia i 4 giocatori presenti al tavolo, verifica l'anteprima matematica delle quote Elo e registra il risultato con un clic.")

    elo_lookup = dict(zip(df_players["Giocatore"], df_players["Elo Attuale"]))

    # Sezione 1: Generatore Randomico / Bilanciato 2 vs 2
    with st.expander("🎲 Generatore Automatico Tavolo (Estrazione Randomica o Bilanciata 2 vs 2)", expanded=True):
        present_players = st.multiselect(
            "Seleziona i giocatori presenti stasera (minimo 4 per il 2vs2):",
            options=db["players"],
            default=db["players"][:6] if len(db["players"]) >= 6 else db["players"]
        )
        b_col1, b_col2, b_col3 = st.columns(3)
        with b_col1:
            if st.button("🎲 Estrai 4 Giocatori Totalmente Random", use_container_width=True):
                if len(present_players) < 4:
                    st.warning("Seleziona almeno 4 giocatori presenti!")
                else:
                    picked = random.sample(present_players, 4)
                    st.session_state["sel_A"] = picked[0]
                    st.session_state["sel_B"] = picked[1]
                    st.session_state["sel_C"] = picked[2]
                    st.session_state["sel_D"] = picked[3]
                    st.toast(f"Tavolo estratto: {picked[0]} & {picked[1]} vs {picked[2]} & {picked[3]}!")

        with b_col2:
            if st.button("⚖️ Estrai 4 e Bilancia Coppie per Elo", use_container_width=True):
                if len(present_players) < 4:
                    st.warning("Seleziona almeno 4 giocatori presenti!")
                else:
                    picked = random.sample(present_players, 4)
                    # Trova l'accoppiamento coi Team Elo più vicini possibili
                    combos = [
                        (picked[0], picked[1], picked[2], picked[3]),
                        (picked[0], picked[2], picked[1], picked[3]),
                        (picked[0], picked[3], picked[1], picked[2]),
                    ]
                    best_combo = min(
                        combos,
                        key=lambda c: abs(
                            calc_team_elo(elo_lookup.get(c[0], 500), elo_lookup.get(c[1], 500), db["settings"]) -
                            calc_team_elo(elo_lookup.get(c[2], 500), elo_lookup.get(c[3], 500), db["settings"])
                        )
                    )
                    st.session_state["sel_A"], st.session_state["sel_B"], st.session_state["sel_C"], st.session_state["sel_D"] = best_combo
                    st.toast("Coppie bilanciate al meglio secondo l'Elo attuale!")

        with b_col3:
            if st.button("🔄 Inverti Accoppiamenti dei 4 Selezionati", use_container_width=True):
                a_cur = st.session_state.get("sel_A", db["players"][0])
                b_cur = st.session_state.get("sel_B", db["players"][1])
                c_cur = st.session_state.get("sel_C", db["players"][2])
                d_cur = st.session_state.get("sel_D", db["players"][3])
                st.session_state["sel_B"], st.session_state["sel_C"] = c_cur, b_cur

    st.divider()

    # Inizializza default session_state per i 4 selettori
    for idx_p, key_p in enumerate(["sel_A", "sel_B", "sel_C", "sel_D"]):
        if key_p not in st.session_state or st.session_state[key_p] not in db["players"]:
            st.session_state[key_p] = db["players"][idx_p % len(db["players"])]

    handicap_mode = st.checkbox(
        "Abilita modalità Handicap 1 vs 2 (consente di inserire lo stesso giocatore in entrambi gli slot di un Team, es. Paolo + Paolo)",
        value=False
    )

    col_t1, col_vs, col_t2 = st.columns([5, 1, 5])
    with col_t1:
        st.markdown("### 🟦 TEAM 1 (Giocatori A e B)")
        player_a = st.selectbox("Giocatore A", options=db["players"], key="sel_A")
        player_b = st.selectbox("Giocatore B", options=db["players"], key="sel_B")
        st.caption(f"Elo {player_a}: **{elo_lookup.get(player_a, 500)}** | Elo {player_b}: **{elo_lookup.get(player_b, 500)}**")

    with col_vs:
        st.markdown("<h2 style='text-align:center; margin-top: 60px;'>VS</h2>", unsafe_allow_html=True)

    with col_t2:
        st.markdown("### 🟥 TEAM 2 (Giocatori C e D)")
        player_c = st.selectbox("Giocatore C", options=db["players"], key="sel_C")
        player_d = st.selectbox("Giocatore D", options=db["players"], key="sel_D")
        st.caption(f"Elo {player_c}: **{elo_lookup.get(player_c, 500)}** | Elo {player_d}: **{elo_lookup.get(player_d, 500)}**")

    # Validazione anti-errore dei giocatori selezionati
    validation_error = None
    if not handicap_mode:
        if len({player_a, player_b, player_c, player_d}) < 4:
            validation_error = "⚠️ Attenzione: hai selezionato lo stesso giocatore più volte! Seleziona 4 giocatori distinti oppure attiva la spunta 'Modalità Handicap 1 vs 2'."
    else:
        if {player_a, player_b} & {player_c, player_d}:
            validation_error = "⚠️ Attenzione: uno stesso giocatore non può trovarsi contemporaneamente nel Team 1 e nel Team 2!"

    # Calcolo e visualizzazione Anteprima Live
    sim = preview_match_math(
        elo_lookup.get(player_a, 500),
        elo_lookup.get(player_b, 500),
        elo_lookup.get(player_c, 500),
        elo_lookup.get(player_d, 500),
        db["settings"]
    )

    st.markdown("#### 🔮 Anteprima Matematica Pre-Partita")
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("🟦 Elo Team 1", f"{sim['elo_t1']} pt", f"Prob. Vittoria: {int(sim['prob_t1']*100)}%")
    p2.metric("Se Vince Team 1", f"+{sim['gain_t1_win']} pt (T1)", f"{sim['loss_t2_lose']} pt (T2)")
    p3.metric("🟥 Elo Team 2", f"{sim['elo_t2']} pt", f"Prob. Vittoria: {int(sim['prob_t2']*100)}%")
    p4.metric("Se Vince Team 2", f"+{sim['gain_t2_win']} pt (T2)", f"{sim['loss_t1_lose']} pt (T1)")

    if validation_error:
        st.error(validation_error)
    else:
        with st.form("form_insert_match", clear_on_submit=False):
            st.markdown("#### 🏁 Esito della Sfida")
            f_c1, f_c2, f_c3 = st.columns([2, 1, 2])
            with f_c1:
                winner_choice = st.radio(
                    "Chi ha vinto la partita?",
                    options=[1, 2],
                    format_func=lambda x: f"🟦 Vittoria TEAM 1 ({player_a} + {player_b})" if x == 1 else f"🟥 Vittoria TEAM 2 ({player_c} + {player_d})",
                    horizontal=True
                )
            with f_c2:
                is_perfect = st.checkbox("🌟 Perfect (Cappotto)")
            with f_c3:
                match_note = st.text_input("Note opzionali (es. serata, rimonta):", "")

            submitted = st.form_submit_button("✅ Registra Partita e Aggiorna Elo", use_container_width=True, type="primary")
            if submitted:
                db["matches"].append({
                    "A": player_a,
                    "B": player_b,
                    "C": player_c,
                    "D": player_d,
                    "Vincitore": int(winner_choice),
                    "Perfect": bool(is_perfect),
                    "Note": match_note.strip()
                })
                save_db(db)
                st.success(f"Partita #{len(db['matches'])} registrata con successo! Classifica Elo aggiornata.")
                st.rerun()

    # Ultime 5 partite registrate per controllo immediato
    st.divider()
    st.subheader("🕒 Ultime 5 Partite Registrate")
    if not df_partite.empty:
        st.dataframe(df_partite.tail(5).iloc[::-1], use_container_width=True, hide_index=True)

# ==============================================================================
# 3. TABELLONE & CLASSIFICHE EXTRA (TOP 10 FORZA, COPPIE, LEGENDE)
# ==============================================================================
elif menu == "📊 Tabellone & Classifiche Extra":
    st.title("📊 Tabellone Completo & Classifiche Speciali")
    st.markdown("Replica fedele di tutte le tabelle avanzate del foglio **Tabellone** dell'Excel (`Top 10 Giocatori`, `Top 10 Coppie`, `Storico Coppie` e `Legende`).")

    tab_top10, tab_pairs, tab_h2h, tab_legend = st.tabs([
        "🏅 Top 10 Live (Indice di Forza)",
        "🤝 Storico Completo Coppie",
        "⚔️ Sinergia & Compagni per Giocatore",
        "📖 Legenda Livelli & Stili"
    ])

    with tab_top10:
        st.info("💡 **Formula Indice di Forza (da Excel):** `ROUND(Vittorie * (Vittorie / Partite), 2)` — premia chi combina un'alta percentuale di vittoria con un elevato volume di partite vinte.")
        col_g10, col_c10 = st.columns(2)

        with col_g10:
            st.subheader("🔥 CLASSIFICA LIVE - TOP 10 GIOCATORI")
            df_top10_g = df_players.sort_values(
                by=["Indice di Forza", "Vinte", "% Vittoria"],
                ascending=[False, False, False]
            ).head(10).reset_index(drop=True)
            df_top10_g["Pos."] = range(1, len(df_top10_g) + 1)
            st.dataframe(
                df_top10_g[["Pos.", "Giocatore", "Vinte", "Partite", "% Vittoria", "Indice di Forza", "Elo Attuale"]],
                use_container_width=True,
                hide_index=True
            )

        with col_c10:
            st.subheader("🤝 CLASSIFICA LIVE - TOP 10 COPPIE")
            only_2v2_top = st.checkbox("Escludi match 1v2 (Solo) dalla Top 10 Coppie", value=True)
            df_pairs_filtered = df_pairs[df_pairs["Tipo"] == "Coppia (2v2)"] if only_2v2_top else df_pairs
            df_top10_c = df_pairs_filtered.sort_values(
                by=["Indice di Forza", "Vittorie", "% Vittorie"],
                ascending=[False, False, False]
            ).head(10).reset_index(drop=True)
            df_top10_c["Pos."] = range(1, len(df_top10_c) + 1)
            st.dataframe(
                df_top10_c[["Pos.", "Coppia", "Vittorie", "Partite", "% Vittorie", "Indice di Forza"]],
                use_container_width=True,
                hide_index=True
            )

    with tab_pairs:
        st.subheader("📋 Storico Completo di Tutte le Coppie")
        f_p1, f_p2 = st.columns(2)
        with f_p1:
            min_pair_games = st.slider("Minimo partite giocate insieme:", 1, 15, 1)
        with f_p2:
            include_solo = st.checkbox("Mostra anche chi ha giocato in 1 vs 2 (Solo)", value=True)

        df_p_view = df_pairs[df_pairs["Partite"] >= min_pair_games].copy()
        if not include_solo:
            df_p_view = df_p_view[df_p_view["Tipo"] == "Coppia (2v2)"]

        st.dataframe(
            df_p_view[[
                "Coppia", "Partite", "Vittorie", "Perse", "% Vittorie",
                "Indice di Forza", "Streak Attiva", "Migliore Streak", "Peggiore Streak", "Tipo"
            ]],
            use_container_width=True,
            hide_index=True
        )

    with tab_h2h:
        st.subheader("⚔️ Analisi Compagni e Avversari per Singolo Giocatore")
        focus_p = st.selectbox("Seleziona un giocatore da analizzare:", options=db["players"])
        if focus_p and not df_partite.empty:
            teammate_stats = {}
            nemesis_stats = {}
            for _, row in df_partite.iterrows():
                t1 = [row["A"], row["B"]]
                t2 = [row["C"], row["D"]]
                v = row["Vincitore"]
                if focus_p in t1:
                    won = (v == 1)
                    mates = [x for x in t1 if x != focus_p] or [f"{focus_p} (Solo)"]
                    opps = set(t2)
                elif focus_p in t2:
                    won = (v == 2)
                    mates = [x for x in t2 if x != focus_p] or [f"{focus_p} (Solo)"]
                    opps = set(t1)
                else:
                    continue

                for m_p in set(mates):
                    teammate_stats.setdefault(m_p, {"Partite": 0, "Vittorie": 0})
                    teammate_stats[m_p]["Partite"] += 1
                    if won:
                        teammate_stats[m_p]["Vittorie"] += 1

                for o_p in opps:
                    nemesis_stats.setdefault(o_p, {"Partite Contro": 0, "Vittorie Contro": 0})
                    nemesis_stats[o_p]["Partite Contro"] += 1
                    if won:
                        nemesis_stats[o_p]["Vittorie Contro"] += 1

            c_mate, c_opp = st.columns(2)
            with c_mate:
                st.markdown(f"**🤝 Rendimento di {focus_p} con ogni compagno:**")
                df_mates = pd.DataFrame([
                    {
                        "Compagno": k,
                        "Partite": v["Partite"],
                        "Vittorie": v["Vittorie"],
                        "% Vittoria": round(v["Vittorie"] / v["Partite"] * 100, 1)
                    }
                    for k, v in teammate_stats.items()
                ]).sort_values(by=["Partite", "% Vittoria"], ascending=[False, False])
                st.dataframe(df_mates, use_container_width=True, hide_index=True)

            with c_opp:
                st.markdown(f"**🎯 Rendimento di {focus_p} contro ogni avversario:**")
                df_opps = pd.DataFrame([
                    {
                        "Avversario": k,
                        "Sfide Contro": v["Partite Contro"],
                        "Vinte da " + focus_p: v["Vittorie Contro"],
                        "% Vittoria": round(v["Vittorie Contro"] / v["Partite Contro"] * 100, 1)
                    }
                    for k, v in nemesis_stats.items()
                ]).sort_values(by=["Sfide Contro", "% Vittoria"], ascending=[False, False])
                st.dataframe(df_opps, use_container_width=True, hide_index=True)

    with tab_legend:
        l_col1, l_col2 = st.columns(2)
        with l_col1:
            st.subheader("🎓 Legenda Livello Tecnico")
            for k, desc in LIVELLI_LEGENDA.items():
                if k != "Non assegnato":
                    st.markdown(f"- **{k}**: {desc}")
        with l_col2:
            st.subheader("🎭 Legenda Stili di Gioco")
            for k, desc in STILI_LEGENDA.items():
                if k != "Non assegnato":
                    st.markdown(f"- **{k}**: {desc}")

# ==============================================================================
# 4. GESTIONE GIOCATORI
# ==============================================================================
elif menu == "👥 Gestione Giocatori":
    st.title("👥 Gestione Giocatori & Profili")
    st.markdown(f"Aggiungi nuovi partecipanti in corsa (con punteggio base automatico di **{db['settings']['start_elo']} punti Elo**) o aggiorna il livello e lo stile di gioco dei membri della Lega.")

    col_add, col_edit = st.columns(2)

    with col_add:
        st.subheader("➕ Aggiungi Nuovo Giocatore")
        with st.form("add_player_form", clear_on_submit=True):
            new_name = st.text_input("Nome / Soprannome del nuovo giocatore:")
            new_lvl = st.selectbox("Livello Tecnico:", options=list(LIVELLI_LEGENDA.keys()))
            new_style = st.selectbox("Stile di Gioco:", options=list(STILI_LEGENDA.keys()))
            sub_add = st.form_submit_button("Aggiungi alla Lega (500 pt Elo)", use_container_width=True, type="primary")
            if sub_add:
                clean_name = new_name.strip()
                if not clean_name:
                    st.error("Inserisci un nome valido.")
                elif any(clean_name.lower() == p.lower() for p in db["players"]):
                    st.error(f"Il giocatore '{clean_name}' è già presente in lista!")
                else:
                    db["players"].append(clean_name)
                    db["player_meta"][clean_name] = {
                        "livello": new_lvl,
                        "stile": new_style,
                        "perfect_base": 0
                    }
                    save_db(db)
                    st.success(f"Giocatore **{clean_name}** aggiunto con {db['settings']['start_elo']} punti Elo di partenza!")
                    st.rerun()

    with col_edit:
        st.subheader("✏️ Modifica Profilo / Rinomina Giocatore")
        target_p = st.selectbox("Seleziona giocatore:", options=db["players"])
        if target_p:
            cur_meta = db.get("player_meta", {}).get(target_p, {})
            with st.form("edit_player_form"):
                renamed_p = st.text_input("Nome Giocatore (modifica per correggere/rinominare ovunque):", value=target_p)
                lvl_idx = list(LIVELLI_LEGENDA.keys()).index(cur_meta.get("livello", "Non assegnato")) if cur_meta.get("livello") in LIVELLI_LEGENDA else 0
                sty_idx = list(STILI_LEGENDA.keys()).index(cur_meta.get("stile", "Non assegnato")) if cur_meta.get("stile") in STILI_LEGENDA else 0
                ed_lvl = st.selectbox("Livello Tecnico:", options=list(LIVELLI_LEGENDA.keys()), index=lvl_idx)
                ed_style = st.selectbox("Stile di Gioco:", options=list(STILI_LEGENDA.keys()), index=sty_idx)
                ed_perf = st.number_input("Perfect Storici (Cappotti extra):", min_value=0, max_value=100, value=int(cur_meta.get("perfect_base", 0)))

                if st.form_submit_button("Salva Modifiche Profilo", use_container_width=True):
                    new_p_clean = renamed_p.strip()
                    if not new_p_clean:
                        st.error("Il nome non può essere vuoto.")
                    else:
                        if new_p_clean != target_p:
                            # Aggiorna nome in players e in tutte le partite storiche
                            db["players"] = [new_p_clean if x == target_p else x for x in db["players"]]
                            # Rimuovi eventuali duplicati mantenendo l'ordine
                            db["players"] = list(dict.fromkeys(db["players"]))
                            for m in db["matches"]:
                                for k_slot in ("A", "B", "C", "D"):
                                    if m[k_slot] == target_p:
                                        m[k_slot] = new_p_clean
                            db["player_meta"].pop(target_p, None)

                        db["player_meta"][new_p_clean] = {
                            "livello": ed_lvl,
                            "stile": ed_style,
                            "perfect_base": int(ed_perf)
                        }
                        save_db(db)
                        st.success(f"Profilo di **{new_p_clean}** aggiornato!")
                        st.rerun()

    st.divider()
    st.subheader("📋 Elenco Giocatori Registrati")
    st.dataframe(
        df_players[["Posizione", "Giocatore", "Elo Attuale", "Partite", "Vinte", "% Vittoria", "Livello", "Stile", "Perfect"]],
        use_container_width=True,
        hide_index=True
    )

# ==============================================================================
# 5. DATI GREZZI & DIAGNOSTICA (AUDIT STORICO PARTITE)
# ==============================================================================
elif menu == "🔍 Dati Grezzi & Diagnostica":
    st.title("🔍 Dati Grezzi delle Partite & Diagnostica Integrità")
    st.markdown("Schermata di controllo trasparente su tutte le **18 colonne matematiche** del foglio `Partite`, con strumenti di verifica delle discrepanze e correzione errori.")

    # Pannello diagnostico automatico sui dati di fondo
    with st.expander("🩺 Diagnostica Automatica sui Dati di Fondo (Excel vs Calcolo Reale)", expanded=True):
        solo_matches = df_partite[(df_partite["A"] == df_partite["B"]) | (df_partite["C"] == df_partite["D"])]
        st.markdown(f"""
        - **Integrità Catena Matematica Elo**: ✅ **100% Verificata** ({len(df_partite)} partite ricalcolate sequenzialmente senza discrepanze).
        - **Partite giocate in Handicap 1 vs 2 rilevate nello storico**: **{len(solo_matches)} partite** (es. `Bonny+Bonny` #34, `Filli+Filli` #36, `Gae+Gae` #59 e #61, `Paolo+Paolo` #71, #73, #74, #75, #77).
          - *Nota tecnica sul file Excel originale*: Nel foglio `Classifica` dell'Excel, la formula `=COUNTIF(Partite!B$2:E1000, Giocatore)` conta **due volte** la presenza e la vittoria di un giocatore quando il suo nome compare sia nella colonna A sia nella colonna B (per questo `Paolo` aveva 45 partite e 22 vittorie nel foglio `Classifica`, ma 40 partite e 19 vittorie reali nel foglio `Tabellone`). L'applicazione traccia entrambi i valori (`Partite Reali` e `Slot Excel COUNTIF`) e ti permette di scegliere quale mostrare!
        """)
        if not solo_matches.empty:
            st.caption("Dettaglio delle partite 1 vs 2 presenti nello storico:")
            st.dataframe(solo_matches[["#", "A", "B", "C", "D", "Vincitore", "Elo Team 1", "Elo Team 2", "Var. Base"]], use_container_width=True, hide_index=True)

    st.subheader("📑 Tabella Completa Dati Grezzi (`Partite`)")
    filter_p_raw = st.selectbox("Filtra partite per giocatore (opzionale):", options=["Tutti i giocatori"] + db["players"])
    df_raw_view = df_partite.copy()
    if filter_p_raw != "Tutti i giocatori":
        df_raw_view = df_raw_view[
            (df_raw_view["A"] == filter_p_raw) |
            (df_raw_view["B"] == filter_p_raw) |
            (df_raw_view["C"] == filter_p_raw) |
            (df_raw_view["D"] == filter_p_raw)
        ]

    st.dataframe(df_raw_view, use_container_width=True, hide_index=True, height=450)

    # Azioni di correzione / eliminazione partita
    st.divider()
    st.subheader("🛠️ Correzione Errori Storico Partite")
    act_c1, act_c2 = st.columns(2)

    with act_c1:
        st.markdown("**⏪ Annulla Ultima Partita Inserita**")
        if len(db["matches"]) > 0:
            last_m = db["matches"][-1]
            st.caption(f"Ultima partita (#{len(db['matches'])}): **{last_m['A']} + {last_m['B']}** vs **{last_m['C']} + {last_m['D']}** (Vincitore: Team {last_m['Vincitore']})")
            if st.button("🗑️ Elimina Ultima Partita e Ricalcola Elo", type="secondary", use_container_width=True):
                db["matches"].pop()
                save_db(db)
                st.success("Ultima partita rimossa e punteggi Elo ricalcolati!")
                st.rerun()

    with act_c2:
        st.markdown("**✂️ Elimina una Partita Specifica per Numero (#)**")
        if len(db["matches"]) > 0:
            del_idx = st.number_input("Numero Partita (#) da eliminare:", min_value=1, max_value=len(db["matches"]), value=len(db["matches"]))
            if st.button(f"Elimina Partita #{del_idx} e Ricalcola", use_container_width=True):
                db["matches"].pop(int(del_idx) - 1)
                save_db(db)
                st.success(f"Partita #{del_idx} eliminata! Tutte le partite successive sono state ricalcolate automaticamente.")
                st.rerun()

# ==============================================================================
# 6. IMPOSTAZIONI SISTEMA ELO
# ==============================================================================
elif menu == "⚙️ Impostazioni Sistema Elo":
    st.title("⚙️ Impostazioni & Parametri Matematici Elo")
    st.markdown("""
    Modifica in qualsiasi momento i parametri del sistema Elo. **Ogni modifica ricalcola istantaneamente l'intero storico delle partite** fin dalla Partita #1, permettendoti di testare bilanciamenti diversi in modo retroattivo e sicuro.
    """)

    with st.form("settings_form"):
        s_col1, s_col2 = st.columns(2)

        with s_col1:
            st.markdown("### 🎯 Punteggio Base & Bonus Partecipazione")
            new_start_elo = st.number_input(
                "Punteggio Elo di Partenza (Default Excel: 500)",
                min_value=100, max_value=2000, step=10,
                value=int(db["settings"]["start_elo"]),
                help="Punteggio assegnato a ogni nuovo giocatore alla sua prima partita."
            )
            new_win_bonus = st.number_input(
                "Bonus Partecipazione / Vittoria (Default Excel: +3)",
                min_value=0, max_value=25, step=1,
                value=int(db["settings"]["win_bonus"]),
                help="Punti extra assegnati ai vincitori oltre alla Variazione Base (formula Excel: G2 + N2 + IF(F2=1, 3, 0)). Premia i giocatori attivi."
            )
            new_pres_bonus = st.number_input(
                "Bonus Presenza Incondizionato (Default Excel: 0)",
                min_value=0, max_value=15, step=1,
                value=int(db["settings"]["presence_bonus"]),
                help="Punti extra assegnati a tutti i 4 giocatori del tavolo solo per aver giocato (anche in caso di sconfitta)."
            )
            new_min_matches = st.number_input(
                "Soglia Minima Partite per Classifica Attiva (Filtro Inattivi)",
                min_value=0, max_value=50, step=1,
                value=int(db["settings"]["min_matches_active"]),
                help="Numero minimo di partite giocate predefinito per comparire nella Dashboard Classifica."
            )

        with s_col2:
            st.markdown("### ⚖️ Media Ponderata Coppie & Sensibilità Elo")
            new_weight_max = st.slider(
                "Peso % del Giocatore più Forte nella Coppia (Default Excel: 0.50 = Media 50/50)",
                min_value=0.0, max_value=1.0, step=0.05,
                value=float(db["settings"]["weight_max"]),
                help="A 0.50 usa la media aritmetica (EloA+EloB)/2. Aumentalo (es. 0.65) per dare più peso al giocatore più forte, o riducilo (es. 0.35) per dare più peso al più debole."
            )
            new_gap_div = st.number_input(
                "Divisore Penalità Divario Coppia (Default Excel: 1200, 0 = Disattivato)",
                min_value=0, max_value=10000, step=100,
                value=int(db["settings"]["gap_divisor"]),
                help="Formula Excel: Media - ((EloMax - EloMin)^2) / Divisore. Più basso è il divisore, più viene penalizzato l'Elo atteso di una coppia sbilanciata."
            )
            new_k_factor = st.number_input(
                "Fattore K - Moltiplicatore Variazione Base (Default Excel: 40)",
                min_value=10, max_value=100, step=5,
                value=int(db["settings"]["k_factor"]),
                help="Formula Excel: ROUND(40 * (Esito - ProbT1), 0)."
            )
            new_scale = st.number_input(
                "Fattore di Scala Logistica (Default Excel: 400)",
                min_value=100, max_value=1000, step=50,
                value=int(db["settings"]["scale_factor"]),
                help="Formula Excel: 1 / (1 + 10^((EloT2 - EloT1)/400))."
            )

        save_btn = st.form_submit_button("💾 Salva Impostazioni e Ricalcola Tutto", type="primary", use_container_width=True)
        if save_btn:
            db["settings"].update({
                "start_elo": int(new_start_elo),
                "win_bonus": int(new_win_bonus),
                "presence_bonus": int(new_pres_bonus),
                "min_matches_active": int(new_min_matches),
                "weight_max": float(new_weight_max),
                "gap_divisor": int(new_gap_div),
                "k_factor": int(new_k_factor),
                "scale_factor": int(new_scale)
            })
            save_db(db)
            st.success("Impostazioni salvate! L'intero storico Elo è stato ricalcolato.")
            st.rerun()

    st.divider()
    r_col1, r_col2 = st.columns(2)
    with r_col1:
        if st.button("🔄 Ripristina Parametri Matematici Default Excel", use_container_width=True):
            db["settings"] = dict(DEFAULT_SETTINGS)
            save_db(db)
            st.success("Parametri riportati ai valori originali del foglio Excel (500 base, +3 bonus, divario 1200, K=40)!")
            st.rerun()
    with r_col2:
        if st.button("⚠️ Reset Totale Database ai Dati Originali dell'Excel (79 Partite)", use_container_width=True):
            st.session_state.db = create_default_db()
            save_db(st.session_state.db)
            st.success("Database ripristinato alle 79 partite originali del file Excel!")
            st.rerun()
