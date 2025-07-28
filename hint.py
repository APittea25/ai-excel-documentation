def generate_individual_hints(summaries: dict) -> dict:
    hint_map = {}

    for name in summaries:
        name_lower = name.lower()
        keywords = []

        if "ax" in name_lower:
            keywords.append("annuity rates")
        if "qx" in name_lower or "mortality" in name_lower:
            keywords.append("mortality rates")
        if "sx" in name_lower:
            keywords.append("survival probabilities")
        if "stoch" in name_lower or "rand" in name_lower or "stochastic" in name_lower:
            keywords.append("simulation-based projections")
        if "vol" in name_lower or "sd" in name_lower or "sigma" in name_lower:
            keywords.append("volatility inputs or stochastic variation")
        if "drift" in name_lower:
            keywords.append("long-term mortality trends or drift terms")
        if "kapp" in name_lower or "beta" in name_lower or "alpha" in name_lower:
            keywords.append("Lee-Carter model parameters")

        if keywords:
            hint_map[name] = "This input/output relates to " + ", ".join(sorted(set(keywords))) + "."

    return hint_map
