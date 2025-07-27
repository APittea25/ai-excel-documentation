def generate_hint_sentence(summaries: dict) -> str:
    hint_keywords = set()

    for name in summaries:
        name_lower = name.lower()
        if "ax" in name_lower:
            hint_keywords.add("annuity rates")
        if "qx" in name_lower or "mortality" in name_lower:
            hint_keywords.add("mortality rates")
        if "sx" in name_lower:
            hint_keywords.add("survival probabilities")
        if "stoch" in name_lower or "rand" in name_lower or "stochastic" in name_lower:
            hint_keywords.add("simulation-based projections")
        if "vol" in name_lower or "sd" in name_lower or "sigma" in name_lower:
            hint_keywords.add("volatility inputs or stochastic variation")
        if "drift" in name_lower:
            hint_keywords.add("long-term mortality trends or drift terms")
        if "kapp" in name_lower or "beta" in name_lower or "alpha" in name_lower:
            hint_keywords.add("Lee-Carter model parameters")

    if hint_keywords:
        return "This model works with " + ", ".join(sorted(hint_keywords)) + "."
    return ""
