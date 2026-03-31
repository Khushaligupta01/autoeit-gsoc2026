"""
AutoEIT GSoC 2026 - Test II: Automated Scoring System
Author: Khushali
Description:
    Applies a meaning-based rubric to EIT transcriptions by comparing
    learner utterances to stimulus sentences. Outputs sentence-level
    scores (0-2) for each participant tab in the Excel file.

Rubric (Meaning-Based EIT Scoring):
    2 - Meaning preserved: Learner conveyed the core meaning of the stimulus,
        even if grammar/vocabulary differs slightly.
    1 - Meaning partially preserved: Some meaningful content reproduced,
        but key elements missing, altered, or unclear.
    0 - Meaning not preserved: Response does not convey the stimulus meaning,
        is unintelligible, mostly gibberish, or empty/xxx.

Approach:
    1. Tokenize both stimulus and learner utterance (after cleaning noise markers).
    2. Compute content word overlap (ignoring function words) as a primary signal.
    3. Use fuzzy string similarity as a secondary signal.
    4. Apply scoring thresholds calibrated to the rubric.
    5. Override to 0 if response is predominantly gibberish/xxx/empty.
"""

import re
import pandas as pd
from rapidfuzz import fuzz

# -------------------------------------------------------------------
# Spanish function words to exclude from content word comparison
# -------------------------------------------------------------------
FUNCTION_WORDS = {
    "el", "la", "los", "las", "un", "una", "unos", "unas",
    "de", "a", "en", "que", "y", "o", "pero", "si", "no",
    "se", "me", "te", "le", "nos", "les", "lo", "con", "por",
    "para", "al", "del", "es", "era", "fue", "ha", "han",
    "su", "sus", "mi", "mis", "tu", "tus", "muy", "mas", "más",
    "que", "como", "cuando", "donde", "quien", "cual",
    "este", "esta", "estos", "estas", "ese", "esa",
    "hay", "ya", "tan", "todo", "toda", "todos", "todas",
    "algo", "nada", "nunca", "siempre", "también"
}

# Noise markers to strip from transcriptions before scoring
NOISE_PATTERN = re.compile(
    r'\[pause\]|\[gibberish\]|\[.*?\]|xxx+|\(.*?\)|\.\.\.|-+',
    re.IGNORECASE
)


def clean_text(text: str) -> str:
    """Remove noise markers, punctuation, normalize whitespace."""
    if not isinstance(text, str):
        return ""
    text = NOISE_PATTERN.sub(" ", text)
    text = re.sub(r"[¿¡!?,;:()\"]", " ", text)
    text = text.lower().strip()
    text = re.sub(r"\s+", " ", text)
    return text


def get_content_words(text: str) -> list:
    """Extract content words (non-function words) from cleaned text."""
    words = clean_text(text).split()
    return [w for w in words if w not in FUNCTION_WORDS and len(w) > 1]


def is_mostly_noise(raw_transcription: str) -> bool:
    """Return True if the response is predominantly gibberish/empty."""
    if not isinstance(raw_transcription, str) or raw_transcription.strip() == "":
        return True
    cleaned = clean_text(raw_transcription)
    if len(cleaned.strip()) < 3:
        return True
    # Count noise tokens in original
    noise_hits = len(re.findall(
        r'\[gibberish\]|xxx+|\[pause\]', raw_transcription, re.IGNORECASE
    ))
    total_tokens = len(raw_transcription.split())
    if total_tokens > 0 and noise_hits / total_tokens > 0.5:
        return True
    return False


def content_word_overlap(stimulus: str, transcription: str) -> float:
    """
    Compute proportion of stimulus content words present in transcription.
    Uses partial matching to handle minor spelling variations.
    """
    stim_words = get_content_words(stimulus)
    trans_words = get_content_words(transcription)

    if not stim_words:
        return 0.0

    matched = 0
    for sw in stim_words:
        # Exact match first
        if sw in trans_words:
            matched += 1
        else:
            # Fuzzy partial match (handles minor misspellings)
            for tw in trans_words:
                if fuzz.ratio(sw, tw) >= 80:
                    matched += 1
                    break

    return matched / len(stim_words)


def score_sentence(stimulus: str, transcription: str) -> int:
    """
    Apply meaning-based rubric to produce a score of 0, 1, or 2.

    Scoring logic:
        - If response is mostly noise/gibberish -> 0
        - Compute content word overlap (0.0 to 1.0)
        - Compute fuzzy full-string similarity (0 to 100)
        - Combine into a weighted score
        - Apply thresholds: >= 0.65 -> 2, >= 0.30 -> 1, else -> 0
    """
    if is_mostly_noise(transcription):
        return 0

    overlap = content_word_overlap(stimulus, transcription)
    fuzzy_sim = fuzz.token_sort_ratio(
        clean_text(stimulus), clean_text(transcription)
    ) / 100.0

    # Weighted combination: content overlap is primary signal
    combined = (0.65 * overlap) + (0.35 * fuzzy_sim)

    if combined >= 0.55:
        return 2
    elif combined >= 0.25:
        return 1
    else:
        return 0


def score_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Apply scoring to a single participant sheet."""
    # Identify stimulus and transcription columns (flexible naming)
    stim_col = None
    trans_col = None

    for col in df.columns:
        col_lower = str(col).lower()
        if "stimulus" in col_lower or "stimul" in col_lower:
            stim_col = col
        if "transcription" in col_lower or "rater" in col_lower:
            trans_col = col

    if stim_col is None or trans_col is None:
        print(f"  Warning: Could not identify columns. Found: {list(df.columns)}")
        return df

    # Apply scorer row by row (skip header/instruction rows)
    scores = []
    for _, row in df.iterrows():
        stimulus = str(row.get(stim_col, ""))
        transcription = str(row.get(trans_col, ""))

        # Skip non-sentence rows
        if not stimulus or stimulus.lower() in ["nan", "stimulus", "sentence", ""]:
            scores.append(None)
            continue

        score = score_sentence(stimulus, transcription)
        scores.append(score)

    # Write scores to Score column
    score_col = None
    for col in df.columns:
        if "score" in str(col).lower():
            score_col = col
            break

    if score_col:
        df[score_col] = scores
    else:
        df["Score"] = scores

    return df


def run_scoring(input_path: str, output_path: str):
    """Main function: read Excel, score all sheets, write output."""
    print(f"\nAutoEIT Automated Scoring System")
    print(f"Input:  {input_path}")
    print(f"Output: {output_path}\n")

    xl = pd.ExcelFile(input_path)
    sheets = xl.sheet_names
    print(f"Found {len(sheets)} sheet(s): {sheets}\n")

    results = {}
    all_scores = []

    for sheet in sheets:
        df = xl.parse(sheet)
        print(f"Processing sheet: '{sheet}' ({len(df)} rows)")

        scored_df = score_sheet(df)
        results[sheet] = scored_df

        # Collect scores for summary
        score_col = None
        for col in scored_df.columns:
            if "score" in str(col).lower():
                score_col = col
                break

        if score_col:
            valid = scored_df[score_col].dropna()
            valid_scores = [s for s in valid if isinstance(s, (int, float))]
            if valid_scores:
                total = sum(valid_scores)
                max_possible = len(valid_scores) * 2
                pct = (total / max_possible) * 100 if max_possible > 0 else 0
                print(f"  Scored {len(valid_scores)} sentences")
                print(f"  Total score: {total}/{max_possible} ({pct:.1f}%)")
                print(f"  Score distribution: "
                      f"2={valid_scores.count(2)}, "
                      f"1={valid_scores.count(1)}, "
                      f"0={valid_scores.count(0)}")
                all_scores.extend(valid_scores)
        print()

    # Write to output Excel
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet, df in results.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

    print(f"Scored file saved to: {output_path}")

    if all_scores:
        print(f"\nOverall Summary across all participants:")
        print(f"  Total sentences scored: {len(all_scores)}")
        print(f"  Overall score: {sum(all_scores)}/{len(all_scores)*2} "
              f"({sum(all_scores)/(len(all_scores)*2)*100:.1f}%)")
        print(f"  Score 2 (full meaning): {all_scores.count(2)} "
              f"({all_scores.count(2)/len(all_scores)*100:.1f}%)")
        print(f"  Score 1 (partial meaning): {all_scores.count(1)} "
              f"({all_scores.count(1)/len(all_scores)*100:.1f}%)")
        print(f"  Score 0 (no meaning): {all_scores.count(0)} "
              f"({all_scores.count(0)/len(all_scores)*100:.1f}%)")


# -------------------------------------------------------------------
# Demo: test with the data from the screenshot
# -------------------------------------------------------------------
if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        output_file = input_file.replace(".xlsx", "_scored.xlsx")
        run_scoring(input_file, output_file)
    else:
        print("Please provide input Excel file path")
        sample_data = [
            ("Quiero cortarme el pelo (7)", "Quiero cortarme mi pelo"),
            ("El libro está en la mesa (7)", "El libro [pause] está en la mesa"),
            ("El carro lo tiene Pedro (8)", "E-[gibberish] perro"),
            ("El se ducha cada mañana (9)", "El se lucha cada mañana"),
            ("¿Qué dice usted que va a hacer hoy? (9)", "¿Qué [gibberish] que vas estoy?"),
            ("Dudo que sepa manejar muy bien (10)", "Dudo que sepa ma-mastar tan bien (tambien?)"),
            ("Las calles de esta ciudad son muy anchas (11)", "Las calles..es-[gibberish]..."),
            ("Puede que llueva mañana todo el día (12)", "Puede xxx mañana de todo día"),
            ("Las casas son muy bonitas pero caras (12)", "A las casa es mu-son bonitas"),
            ("Me gustan las películas que acaban bien (12)", "Me gusta las películas que x bien"),
            ("El chico con el que yo salgo es español (13)", "El chico con es el algo (xxx?) es español"),
            ("Después de cenar me fui a dormir tranquilo (13)", "Después de fenar mi fui a tranquilo"),
            ("Quiero una casa en la que vivan mis animales (14)", "Quienen que en el casa ...muchos animales"),
            ("A nosotros nos fascinan las fiestas grandiosas (14)", "A nosotros fanimos [pause] fiestas grandiosas"),
            ("Ella sólo bebe cerveza y no come nada (15)", "Ella bebidas cerviamos y no comidas"),
            ("Me gustaría que el precio de las casas bajara (15)", "Me gustaría ser..ella...mhh.."),
            ("Cruza a la derecha y después sigue todo recto (15)", "Cruze a la derecha de siguelo"),
            ("Ella ha terminado de pintar su apartamento (15)", "Ella terminado pintar a sus apartmento"),
            ("Me gustaría que empezara a hacer más calor pronto (15)", "Me gustaría se..[pause] xxx pronto"),
            ("El niño al que se le murió el gato está triste (16)", "El niño se murió el gato es muy triste"),
            ("Una amiga mía cuida a los niños de mi vecino (16)", "Una mía amiga cuidado a sus niños"),
            ("El gato que era negro fue perseguido por el perro (16)", "El gato-el gato quien era el nego de perro"),
            ("Antes de poder salir él tiene que limpiar su cuarto (16)", "Antes de podai e salir..[pause] antes de xxx"),
            ("La cantidad de personas que fuman ha disminuido (17)", "A la cantan... [pause] muy a xxx"),
            ("Después de llegar a casa del trabajo tomé la cena (17)", "Después de llegar [pause] al trabajo..."),
            ("El ladrón al que atrapó la policía era famoso (17)", "X- ladróna(?)de policio que es famoso"),
            ("Le pedí a un amigo que me ayudara con la tarea (17)", "Que ped-x una amigo [pause] a su tarea"),
            ("El examen no fue tan difícil como me habían dicho (17)", "El examen no difícil.. [pause]... [gibberish]"),
            ("¿Serías tan amable de darme el libro que está en la mesa? (17)", "El libro que está en la mesa"),
            ("Hay mucha gente que no toma nada para el desayuno (17)", "A mucha gente xxx..."),
        ]

        print(f"{'#':<4} {'Score':<6} {'Stimulus (trimmed)':<45} {'Transcription (trimmed)'}")
        print("-" * 100)

        scores = []
        for i, (stim, trans) in enumerate(sample_data, 1):
            s = score_sentence(stim, trans)
            scores.append(s)
            stim_short = stim[:43] + ".." if len(stim) > 43 else stim
            trans_short = trans[:40] + ".." if len(trans) > 40 else trans
            print(f"{i:<4} {s:<6} {stim_short:<45} {trans_short}")

        total = sum(scores)
        max_p = len(scores) * 2
        print(f"\nTotal: {total}/{max_p} ({total/max_p*100:.1f}%)")
        print(f"Score 2: {scores.count(2)} | Score 1: {scores.count(1)} | Score 0: {scores.count(0)}")