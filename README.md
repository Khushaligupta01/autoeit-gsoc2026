# AutoEIT GSoC 2026 — Automated Scoring System

## Overview
This project implements an automated scoring system for Elicited Imitation Tasks (EIT), as part of the HumanAI GSoC 2026 evaluation.

The system evaluates learner transcriptions against stimulus sentences using a meaning-based rubric and outputs sentence-level scores.

---

## Features

- Meaning-based scoring (0, 1, 2)
- Content word overlap analysis
- Fuzzy matching for robustness to variations
- Noise handling (pause, gibberish, incomplete speech)
- Works across multiple participant sheets in Excel

---

## Scoring Rubric

- **2** — Meaning preserved  
- **1** — Meaning partially preserved  
- **0** — Meaning not preserved  

---

## Approach

1. Clean transcription data (remove noise markers like `[pause]`, `xxx`)
2. Extract content words (ignore function words)
3. Compute:
   - Content word overlap
   - Fuzzy similarity (RapidFuzz)
4. Combine both signals into a weighted score
5. Apply threshold-based classification

---

## How to Run

```bash
pip install pandas openpyxl rapidfuzz
python autoeit_scorer.py "AutoEIT Sample Transcriptions for Scoring.xlsx"
