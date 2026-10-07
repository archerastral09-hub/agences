"""Récupère les agences immobilières actives (APE 68.31Z) par département
via l'API publique Recherche d'entreprises (sans clé) et écrit un CSV par
département dans data/agences_XX.csv.

Usage : python fetch_sirene.py all                (les 101 départements)
        python fetch_sirene.py 75                 (un département)
        python fetch_sirene.py 01,02,75           (une liste)
        python fetch_sirene.py depuis:40          (reprend à partir du 40)
        python fetch_sirene.py all --push         (commit + push après chaque département)
        python fetch_sirene.py all --force        (refait aussi les départements déjà présents)

Reprise : un département dont le CSV existe déjà est sauté (sauf --force).

Une société n'est gardée que dans le département de son SIÈGE : l'API renvoie
aussi les sociétés qui n'ont qu'un établissement dans le département, ce qui
créerait des doublons entre fichiers.
"""
import csv
import json
import os
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date

API = "https://recherche-entreprises.api.gouv.fr/search"
NAF = "68.31Z"
PER_PAGE = 25          # maximum autorisé par l'API
PAUSE = 0.25           # l'API limite à 7 appels/seconde
MAX_PAGES = 400        # l'API ne renvoie pas plus de 10 000 résultats

DEPARTEMENTS = (
    [f"{i:02d}" for i in range(1, 96) if i != 20]
    + ["2A", "2B", "971", "972", "973", "974", "976"]
)

COLONNES = ["siren", "siret_siege", "nom", "enseigne", "adresse",
            "code_postal", "ville", "date_creation", "gerants",
            "source", "date_collecte"]

FORCE = False


def get(params):
    url = API + "?" + urllib.parse.urlencode(params)
    for essai in range(6):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "gh-agences/1.0"})
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503, 504):
                time.sleep(3 * (essai + 1))
                continue
            raise
        except (urllib.error.URLError, TimeoutError, ConnectionError,
                json.JSONDecodeError):
            time.sleep(3 * (essai + 1))
    raise RuntimeError("API injoignable : " + url)


def gerants(entreprise):
    noms = []
    for d in entreprise.get("dirigeants") or []:
        # l'API renvoie parfois des champs à null : on les traite comme vides
        nom_famille = d.get("nom") or ""
        if nom_famille:
            prenoms = d.get("prenoms") or ""
            nom = f"{prenoms.title()} {nom_famille.title()}".strip()
        else:
            nom = d.get("denomination") or ""
        if nom:
            q = d.get("qualite") or ""
            noms.append(f"{nom} ({q})" if q else nom)
    return " | ".join(noms)


def ligne_de(e, s):
    enseignes = [x for x in (s.get("liste_enseignes") or []) if x]
    siren = e.get("siren") or ""
    return {
        "siren": siren,
        "siret_siege": s.get("siret") or "",
        "nom": e.get("nom_complet") or "",
        "enseigne": " | ".join(enseignes),
        "adresse": s.get("adresse") or "",
        "code_postal": s.get("code_postal") or "",
        "ville": s.get("libelle_commune") or "",
        "date_creation": e.get("date_creation") or "",
        "gerants": gerants(e),
        "source": "https://annuaire-entreprises.data.gouv.fr/entreprise/" + siren,
        "date_collecte": date.today().isoformat(),
    }


def run(dep):
    sortie = f"data/agences_{dep}.csv"
    if not FORCE and os.path.exists(sortie) and os.path.getsize(sortie) > 0:
        print(f"{dep} : déjà fait, ignoré ({sortie})")
        return False

    brut, page = [], 1
    while page <= MAX_PAGES:
        data = get({"activite_principale": NAF, "departement": dep,
                    "etat_administratif": "A",
                    "per_page": PER_PAGE, "page": page})
        resultats = data.get("results") or []
        if not resultats:
            break
        brut.extend(resultats)
        if page >= (data.get("total_pages") or 1):
            break
        page += 1
        time.sleep(PAUSE)

    if len(brut) >= PER_PAGE * MAX_PAGES:
        print(f"ATTENTION : {dep} atteint la limite de 10 000 résultats, "
              "il faut découper par code postal ou commune.")

    # on ne garde que les sociétés dont le siège est dans ce département
    gardes, autres_dep = [], 0
    for e in brut:
        s = e.get("siege") or {}
        d = (s.get("departement") or "").strip().upper()
        if d and d != dep:
            autres_dep += 1
        else:
            gardes.append((e, s))
    if brut and not gardes:
        # champ « departement » inutilisable : on garde tout plutôt que rien perdre
        print(f"ATTENTION : {dep} aucun siège reconnu dans le département, "
              "tout est gardé (doublons possibles).")
        gardes = [(e, e.get("siege") or {}) for e in brut]
        autres_dep = 0

    vus, lignes = set(), []
    for e, s in gardes:
        ligne = ligne_de(e, s)
        if ligne["siren"] and ligne["siren"] in vus:   # doublon de pagination
            continue
        vus.add(ligne["siren"])
        lignes.append(ligne)

    os.makedirs("data", exist_ok=True)
    tmp = sortie + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLONNES)
        w.writeheader()
        w.writerows(lignes)
    os.replace(tmp, sortie)   # écriture atomique : pas de CSV à moitié écrit
    print(f"{dep} : {len(lignes)} agences écrites dans {sortie} "
          f"({autres_dep} ignorées, siège dans un autre département)")
    return True


def git_push(dep):
    """Commit + push du CSV du département (utilisé par le workflow)."""
    subprocess.run(["git", "add", "data/"], check=True)
    if subprocess.run(["git", "diff", "--cached", "--quiet"]).returncode == 0:
        return
    subprocess.run(["git", "commit", "-m", f"Agences département {dep}"], check=True)
    for essai in range(4):
        if subprocess.run(["git", "push"]).returncode == 0:
            return
        subprocess.run(["git", "pull", "--rebase"])
        time.sleep(2 * (essai + 1))
    raise RuntimeError(f"push impossible pour {dep}")


def norm(c):
    c = c.strip().upper()
    return c.zfill(2) if c.isdigit() and len(c) < 3 else c


def resoudre(texte):
    """all | 75 | 01,02,75 | depuis:40  ->  liste de codes valides."""
    texte = texte.strip().upper()
    if texte == "ALL":
        return list(DEPARTEMENTS)
    if texte.startswith("DEPUIS:"):
        debut = norm(texte.split(":", 1)[1])
        if debut not in DEPARTEMENTS:
            sys.exit(f"Département de départ invalide : {debut}")
        return DEPARTEMENTS[DEPARTEMENTS.index(debut):]
    codes = [norm(c) for c in texte.split(",") if c.strip()]
    invalides = [c for c in codes if c not in DEPARTEMENTS]
    if invalides or not codes:
        sys.exit("Code(s) invalide(s) : " + (", ".join(invalides) or texte) +
                 "\nValides : 01 à 95 (sans 20), 2A, 2B, 971, 972, 973, 974, 976")
    return codes


def main():
    global FORCE
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    FORCE = "--force" in sys.argv
    push = "--push" in sys.argv
    if not args:
        sys.exit(__doc__)
    deps = resoudre(args[0])

    echecs = []
    for dep in deps:
        try:
            if run(dep) and push:
                git_push(dep)
        except Exception as e:   # on continue avec le département suivant
            print(f"ERREUR {dep} : {e}")
            echecs.append(dep)
    if echecs:
        print("Départements en échec (relancer pour reprendre) : " + ", ".join(echecs))
        sys.exit(1)


if __name__ == "__main__":
    main()
