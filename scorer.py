import pandas as pd
import numpy as np
from difflib import SequenceMatcher
import re

# Cache rubric to avoid repeated loading
RUBRIC_CACHE = None


# --------------------------------------------------
# LOAD RUBRIC
# --------------------------------------------------
def read_rubric_from_excel(path):
    global RUBRIC_CACHE

    if RUBRIC_CACHE is not None:
        return RUBRIC_CACHE

    # Default fallback rubric
    fallback = {
        "Content Quality": {
            "keywords": ["introduction", "background", "experience", "skills", "education"],
            "min_words": 40,
            "max_words": 250,
            "weight": 1.0
        },
        "Clarity": {
            "keywords": ["clear", "confident", "communication", "organized"],
            "min_words": 40,
            "max_words": 250,
            "weight": 1.0
        },
        "Structure": {
            "keywords": ["name", "education", "interest", "goal"],
            "min_words": 40,
            "max_words": 250,
            "weight": 1.0
        }
    }

    if path is None:
        RUBRIC_CACHE = fallback
        return fallback

    try:
        df = pd.read_excel(path)

        rubric = {}
        for _, row in df.iterrows():
            name = str(row["Criterion"]).strip()
            keywords = str(row["Keywords"]).split(",")
            min_w = int(row["Min Words"])
            max_w = int(row["Max Words"])
            weight = float(row["Weight"])

            rubric[name] = {
                "keywords": [k.strip().lower() for k in keywords if k.strip()],
                "min_words": min_w,
                "max_words": max_w,
                "weight": weight
            }

        RUBRIC_CACHE = rubric
        return rubric

    except:
        RUBRIC_CACHE = fallback
        return fallback


# --------------------------------------------------
# RULE-BASED SCORING
# --------------------------------------------------
def compute_rule_score(text, keywords, min_words, max_words):
    text_lower = text.lower()
    words = text_lower.split()
    word_count = len(words)

    # Keyword detection
    found = []
    for k in keywords:
        if k in text_lower:
            found.append(k)

    keyword_score = len(found) / len(keywords) if keywords else 1

    # Soft length scoring logic
    if word_count < min_words:
        length_score = 0.6
    elif word_count > max_words:
        length_score = 0.7
    else:
        length_score = 1.0

    rule_score = (keyword_score + length_score) / 2
    return rule_score, found, word_count


# --------------------------------------------------
# SEMANTIC SCORING (Improved Meaning-Matching)
# --------------------------------------------------
def compute_semantic_score(text, keywords):
    if not keywords:
        return 0.5

    text = text.lower()

    semantic_scores = []
    for k in keywords:

        # Partial match scoring (experience → experienced, working)
        partial_match = any(word.startswith(k[:4]) for word in text.split())

        if partial_match:
            semantic_scores.append(0.6)
            continue

        # Fallback similarity using SequenceMatcher
        sim = SequenceMatcher(None, k, text).ratio()

        # Reduce noise
        if sim > 0.35:
            semantic_scores.append(sim)
        else:
            semantic_scores.append(0)

    # Avoid harsh zeros
    avg = sum(semantic_scores) / len(keywords)

    # Smooth scaling
    return round(min(1.0, max(0.2, avg)), 2)


# --------------------------------------------------
# FINAL SCORING
# --------------------------------------------------
def score_transcript(text, rubric):
    results = {}
    total_weight = sum(item["weight"] for item in rubric.values())
    weighted_sum = 0

    for crit, info in rubric.items():

        rule_score, found_kw, wc = compute_rule_score(
            text,
            info["keywords"],
            info["min_words"],
            info["max_words"]
        )

        semantic_score = compute_semantic_score(text, info["keywords"])

        final_score = ((rule_score + semantic_score) / 2) * 100
        weighted_sum += final_score * info["weight"]

        results[crit] = {
            "score": round(final_score, 2),
            "rule_score": round(rule_score, 2),
            "semantic_score": round(semantic_score, 2),
            "keywords_found": found_kw,
            "num_words": wc,
            "weight": info["weight"]
        }

    overall = weighted_sum / total_weight

    return {
        "overall_score": round(overall, 2),
        "criterion_scores": results
    }

