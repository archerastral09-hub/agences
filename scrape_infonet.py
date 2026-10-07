"""Enrichit les agences (data/agences_XX.csv, produits par fetch_sirene.py)
avec les infos de la fiche infonet.fr : dirigeant, téléphone, email, site, etc.

URL visée : https://infonet.fr/entreprises/<siret>-<nom-slugifié>/

Usage :
  python scrape_infonet.py                       toutes les agences de data/
  python scrape_infonet.py --dep 33,75           certains départements
  python scrape_infonet.py --limit 20            les 20 premières à faire
  python scrape_infonet.py --siret 45199225900015 --nom "ARTEMIA IMMOBILIER"
  python scrape_infonet.py --html fiche.html     parse une page enregistrée (aucun réseau)
  python scrape_infonet.py --navigateur          rendu JavaScript via Playwright

Reprise automatique : data/infonet.csv est écrit ligne par ligne. Au relancement,
les SIRET déjà présents (statut ok ou introuvable) sont sautés. Les erreurs sont
journalisées dans data/infonet_erreurs.csv et retentées au prochain passage.
Ctrl+C ou plantage : rien n'est perdu, il suffit de relancer la même commande.
"""
import argparse
import csv
import glob
import os
import random
import re
import subprocess
import sys
import time
import unicodedata
from datetime import datetime

import requests
from bs4 import BeautifulSoup

BASE = "https://infonet.fr/entreprises/"
SORTIE = "data/infonet.csv"
ERREURS = "data/infonet_erreurs.csv"
DELAI = (1.5, 3.0)        # pause aléatoire entre deux fiches (secondes)
MAX_ECHECS_DE_SUITE = 5   # arrêt propre si le site bloque / change

CHAMPS = [
    ("dirigeant", "Dirigeant"), ("telephone", "Téléphone"), ("email", "Email"),
    ("site_web", "Site web"), ("linkedin", "LinkedIn"), ("x", "X"),
    ("greffe", "Greffe"), ("rcs", "RCS"), ("siren", "SIREN"), ("siret", "SIRET"),
    ("tva", "N° TVA"), ("conv_collectives", "Conv. collectives"),
    ("diffusion", "Diffusion"), ("taille", "Taille entreprise"),
    ("etablissement", "Établissement"), ("type_exercice", "Type d'exercice"),
    ("code_ape", "Code APE"), ("type_activite", "Type activité"),
    ("forme_juridique", "Forme juridique"), ("capital_social", "Capital social"),
    ("date_creation", "Date de création"), ("maj_fiche", "Fiche mise à jour"),
    ("cloture_exercice", "Clôture exercice"),
]
COLONNES = ["siret_entree", "nom_entree", "url", "statut", "date_collecte"] + \
           [c for c, _ in CHAMPS]


class Bloque(Exception):
    """403 / 429 / 5xx persistants : le site refuse ou est indisponible."""


def norm(t):
    t = (t or "").replace("\xa0", " ").replace("’", "'").strip().rstrip(":").strip()
    return unicodedata.normalize("NFC", t).casefold()


LABELS = {norm(lib): cle for cle, lib in CHAMPS}


def slug(nom):
    s = unicodedata.normalize("NFKD", nom or "").encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def parse(html):
    """Extrait les champs d'une fiche. Retourne {cle: valeur} (clés trouvées seulement)."""
    soup = BeautifulSoup(html, "html.parser")
    for t in soup(["script", "style", "noscript"]):
        t.decompose()
    lignes = [l.strip() for l in soup.get_text("\n").split("\n") if l.strip()]
    out = {}
    for i, l in enumerate(lignes):
        cle = LABELS.get(norm(l))
        valeur = None
        if cle is None and ":" in l:                      # « Label : valeur » sur une ligne
            debut, _, reste = l.partition(":")
            cle = LABELS.get(norm(debut))
            valeur = reste.strip() if cle else None
        elif cle is not None:                             # label seul, valeur à la ligne suivante
            suivante = lignes[i + 1] if i + 1 < len(lignes) else ""
            valeur = "" if norm(suivante) in LABELS else suivante
        if cle and cle not in out and valeur is not None:
            out[cle] = valeur
    # secours via les liens cliquables si le texte est vide ou masqué
    for cle, prefixe in (("email", "mailto:"), ("telephone", "tel:")):
        if not out.get(cle):
            a = soup.select_one(f'a[href^="{prefixe}"]')
            if a:
                out[cle] = a["href"][len(prefixe):].strip()
    return out


# ---------- récupération de la page ----------
_session = requests.Session()
_session.headers.update({
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "fr-FR,fr;q=0.9",
})
_pw = {}


def fetch_http(url):
    for essai in range(5):
        try:
            r = _session.get(url, timeout=30)
        except requests.RequestException:
            time.sleep(3 * (essai + 1))
            continue
        if r.status_code == 200:
            return 200, r.text
        if r.status_code == 404:
            return 404, ""
        if r.status_code in (403, 429, 500, 502, 503, 504):
            attente = int(r.headers.get("Retry-After", 0) or 0) or 10 * (essai + 1)
            time.sleep(min(attente, 120))
            continue
        return r.status_code, ""
    raise Bloque(f"échec réseau/blocage persistant : {url}")


def fetch_navigateur(url):
    if "page" not in _pw:
        from playwright.sync_api import sync_playwright
        _pw["p"] = sync_playwright().start()
        _pw["b"] = _pw["p"].chromium.launch()
        _pw["page"] = _pw["b"].new_page(locale="fr-FR")
    page = _pw["page"]
    rep = page.goto(url, wait_until="networkidle", timeout=45000)
    code = rep.status if rep else 0
    if code in (403, 429, 503):
        raise Bloque(f"HTTP {code} : {url}")
    return code, (page.content() if code == 200 else "")


def collecter(siret, nom, navigateur):
    fetch = fetch_navigateur if navigateur else fetch_http
    urls = [f"{BASE}{siret}-{slug(nom)}/", f"{BASE}{siret}/"]   # 2e essai sans slug
    for url in urls:
        code, html = fetch(url)
        if code == 200:
            champs = parse(html)
            return url, ("ok" if len(champs) >= 5 else "vide"), champs
        if code != 404:
            raise Bloque(f"HTTP {code} : {url}")
    return urls[0], "introuvable", {}


# ---------- entrées / sorties / reprise ----------
def lire_agences(deps):
    vus = set()
    for chemin in sorted(glob.glob("data/agences_*.csv")):
        dep = os.path.basename(chemin)[8:-4]
        if deps and dep not in deps:
            continue
        with open(chemin, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                siret = (row.get("siret_siege") or "").strip()
                if siret and siret not in vus:
                    vus.add(siret)
                    yield siret, row.get("nom") or ""


def deja_faits():
    faits = set()
    if os.path.exists(SORTIE):
        with open(SORTIE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("statut") in ("ok", "introuvable"):
                    faits.add(row["siret_entree"])
    return faits


def ajouter(chemin, colonnes, row):
    nouveau = not os.path.exists(chemin) or os.path.getsize(chemin) == 0
    os.makedirs(os.path.dirname(chemin), exist_ok=True)
    with open(chemin, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=colonnes, extrasaction="ignore")
        if nouveau:
            w.writeheader()
        w.writerow(row)
        f.flush()
        os.fsync(f.fileno())      # ligne sur disque avant de passer à la suivante


def git_push(message):
    """Commit + push de data/ (reprise : l'état vit dans le dépôt)."""
    subprocess.run(["git", "add", "data/"], check=True)
    if subprocess.run(["git", "diff", "--cached", "--quiet"]).returncode == 0:
        return
    subprocess.run(["git", "commit", "-m", message], check=True)
    for essai in range(4):
        if subprocess.run(["git", "push"]).returncode == 0:
            return
        subprocess.run(["git", "pull", "--rebase", "--autostash"])
        time.sleep(2 * (essai + 1))
    raise RuntimeError("push impossible")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dep", default="")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--siret")
    ap.add_argument("--nom", default="")
    ap.add_argument("--html")
    ap.add_argument("--navigateur", action="store_true")
    ap.add_argument("--push", action="store_true",
                    help="commit + push de data/ régulièrement et à la fin")
    ap.add_argument("--push-tous", type=int, default=100,
                    help="avec --push : un commit toutes les N fiches (défaut 100)")
    ap.add_argument("--duree-max", type=int, default=0,
                    help="arrêt propre après N minutes (0 = illimité)")
    a = ap.parse_args()
    debut = time.time()

    if a.html:
        with open(a.html, encoding="utf-8") as f:
            for k, v in parse(f.read()).items():
                print(f"{k:18} {v}")
        return

    if a.siret:
        taches = [(a.siret, a.nom)]
        faits = set()
    else:
        deps = {d.strip().upper().zfill(2) if d.strip().isdigit() and len(d.strip()) < 3
                else d.strip().upper() for d in a.dep.split(",") if d.strip()}
        faits = deja_faits()
        taches = [t for t in lire_agences(deps) if t[0] not in faits]
    if a.limit:
        taches = taches[:a.limit]
    print(f"{len(faits)} déjà faites, {len(taches)} à traiter")

    echecs = ok = 0
    try:
        for n, (siret, nom) in enumerate(taches, 1):
            if a.duree_max and time.time() - debut > a.duree_max * 60:
                print(f"Durée max ({a.duree_max} min) atteinte : arrêt propre, "
                      "relance le workflow pour continuer.")
                break
            ligne = {"siret_entree": siret, "nom_entree": nom,
                     "date_collecte": datetime.now().isoformat(timespec="seconds")}
            try:
                url, statut, champs = collecter(siret, nom, a.navigateur)
                ligne.update(champs)
                ligne.update(url=url, statut=statut)
                if statut == "vide":
                    raise ValueError("page sans champs reconnus (rendu JS ? --navigateur)")
                ajouter(SORTIE, COLONNES, ligne)
                ok += 1
                echecs = 0
                print(f"[{n}/{len(taches)}] {siret} {statut}")
                if a.push and ok % a.push_tous == 0:
                    git_push(f"Infonet : {ok} fiches")
            except Bloque as e:
                ajouter(ERREURS, ["date", "siret", "nom", "erreur"],
                        {"date": ligne["date_collecte"], "siret": siret, "nom": nom, "erreur": str(e)})
                echecs += 1
                print(f"[{n}/{len(taches)}] {siret} BLOQUÉ : {e}")
            except Exception as e:
                ajouter(ERREURS, ["date", "siret", "nom", "erreur"],
                        {"date": ligne["date_collecte"], "siret": siret, "nom": nom,
                         "erreur": f"{type(e).__name__}: {e}"})
                echecs += 1
                print(f"[{n}/{len(taches)}] {siret} ERREUR : {type(e).__name__}: {e}")
            if echecs >= MAX_ECHECS_DE_SUITE:
                print(f"{echecs} échecs de suite : arrêt. Relance la même commande plus tard.")
                sys.exit(2)
            time.sleep(random.uniform(*DELAI))
    except KeyboardInterrupt:
        print("\nInterrompu : les lignes déjà écrites sont conservées, relance pour reprendre.")
    finally:
        if a.push:                      # même après un blocage (sys.exit) ou une erreur
            git_push(f"Infonet : {ok} fiches (fin de passage)")
    print(f"Terminé : {ok} fiches enregistrées.")


if __name__ == "__main__":
    main()
