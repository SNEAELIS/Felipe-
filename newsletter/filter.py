import re
from typing import Dict, List, Any

# Define keyword sets tailored for Brazilian politics
CATEGORY_KEYWORDS: Dict[str, List[str]] = {
    "Economia": [
        "PIB", "inflação", "Selic", "juros", "imposto", "reforma tributária",
        "ministério da fazenda", "orçamento", "câmbio", "dólar", "B3", "mercado",
        "haddad", "arcabouço"
    ],
    "Justiça / Policial": [
        "STF", "PF", "polícia federal", "investigação", "processo", "habeas corpus",
        "denúncia", "PGR", "inquérito", "condenação", "justiça", "tribunal", 
        "operação", "moraes", "voter", "prisão"
    ],
    "Política / Legislativo": [
        "Câmara", "Senado", "votação", "projeto de lei", "PL", "medida provisória",
        "partido", "eleições", "relator", "comissão", "governo", "oposição",
        "congresso", "deputado", "senador"
    ],
    "Bastidores / Redes Sociais": [
        "gafe", "bastidores", "instagram", "post", "compara", "polêmica",
        "meme", "viagem", "susto", "flagrado", "x", "twitter"
    ]
}

def classify_title(title: str) -> str:
    """Classifies a single news title based on keyword frequency."""
    title_lower = title.lower()
    scores = {}

    for category, keywords in CATEGORY_KEYWORDS.items():
        # Count exact word matches in the title
        score = sum(
            1 for kw in keywords 
            if re.search(r'\b' + re.escape(kw) + r'\b', title_lower)
        )
        if score > 0:
            scores[category] = score

    if not scores:
        return "Geral"

    # Return the category with the highest match score
    return max(scores, key=scores.get)


def process_and_group_articles(articles: List[Dict[str, str]]) -> Dict[str, List[Dict[str, str]]]:
    """
    Takes a flat list of article dictionaries and returns them 
    grouped by category for Jinja2 rendering.
    """
    grouped = {
        "Economia": [],
        "Justiça / Policial": [],
        "Política / Legislativo": [],
        "Bastidores / Redes Sociais": [],
        "Geral": []
    }

    for article in articles:
        category = classify_title(article["title"])
        article["category"] = category
        grouped[category].append(article)

    return grouped