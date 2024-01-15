import httpx
import time
import xlsxwriter

from selectolax.parser import HTMLParser
from dataclasses import dataclass , asdict
import pandas as pd


@dataclass
class Webtoon:
    Nom : str | None
    Auteur : str | None
    Genre : str | None
    Vues_Par_Millions : int | None
    Note : float | None
    Follow : int | None

def get_html(base_url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/9.0.4472.164 Safari/537.36"
    }

    resp = httpx.get(base_url, headers=headers)
    html = HTMLParser(resp.text)

    print(html.css_first('h1').text())
    return html  # Retourner la valeur de html

def parse_url(html):
    produits = html.css('a[href]')
    nb = 0
    urls = []

    liens_a_exclure = [
        "https://www.webtoons.com/en/terms",
        "https://www.webtoons.com/en/terms/privacyPolicy",
        "https://www.webtoons.com/en/consentsManagement",
        "https://www.webtoons.com/en/advertising",
        "https://www.webtoons.com/en/terms/dnsmpi",
        "https://www.webtoons.com/en/contact",
        'https://www.webtoons.com/en/',
        'https://www.webtoons.com/en/creators101/webtoon-canvas',
        'https://www.webtoons.com/en/originals',
        'https://www.webtoons.com/en/genres',
        'https://www.webtoons.com/en/popular',
        'https://www.webtoons.com/en/canvas'
    ]

    for produit in produits:
        element = produit.css_first('a[href]')
        if element:
            href = element.attributes.get('href')
            if href and href != '#' and href.startswith("https://www.webtoons.com/en/") and href not in liens_a_exclure:
                #print(href)
                #time.sleep(0.1)
                urls.append(href)
                #nb += 1
    #print(f"Nombre de liens trouvés : {nb}")
    #print("Liste des liens :", urls)
    #print(f"Nombre d'éléments dans la liste : {len(urls)}")
    return urls  # Retourner la liste des liens

def clean_author_text(text):
    # Supprimer les caractères indésirables dans le texte de l'auteur
    cleaned_text = text.replace('\n', '').replace('\t', '').replace('author info', '').strip()
    return cleaned_text

def extract_data(html,sel) : 
    try : 
        data = html.css_first(sel).text().strip()
        return data
    except AttributeError : 
        return None

def parse_details_page(html):
    nom = extract_data(html, 'h1')
    
    auteur = extract_data(html, 'a.author._gaLoggingLink')
    if auteur is None:
        auteur = extract_data(html, 'div.author_area')
    
    # Nettoyer le texte de l'auteur si présent
    if auteur:
        auteur = clean_author_text(auteur)

    genre = extract_data(html, 'h2')
    vues_par_millions = extract_data(html, 'em.cnt')
    vues_par_millions = normalize_views(vues_par_millions)
    note = extract_data(html, 'em.cnt#_starScoreAverage')
    follow = extract_data(html, 'li span.ico_subscribe + em.cnt')

    new_webtoon = Webtoon(
        Nom=nom,
        Auteur=auteur,
        Genre=genre,
        Vues_Par_Millions=vues_par_millions,
        Note=note,
        Follow=follow,
    )
    return asdict(new_webtoon)

def normalize_views(views_text):
    if views_text:
        if 'B' in views_text:
            return float(views_text.replace('B', '')) * 1e3  # Convertir en millions (1 B = 1e3 M)
        elif 'M' in views_text:
            return float(views_text.replace('M', ''))
    return None


def clean_data(webtoons):
    df = pd.DataFrame(columns=webtoons[0].keys())
    print(df.columns)
    for webtoon in webtoons:
        length = len(df)
        df.loc[length] = webtoon

    df = df.sort_values(by=['Vues_Par_Millions'], ascending=False)
    return df

def export_to_excel(df):
    writer = pd.ExcelWriter("webtoons.xlsx", engine='xlsxwriter')

    df.to_excel(writer, sheet_name='Sheet1', index=False)
    worksheet = writer.sheets['Sheet1']

    for i, col in enumerate(df.columns):
        max_len = df[col].astype(str).apply(len).max()
        worksheet.set_column(i, i, max_len + 2)  # Ajouter une marge de 2 caractères pour la largeur de la colonne
        worksheet.set_column('D:D', 17)  # Vues_Par_Millions
        worksheet.set_column('E:E', 10)  # Note
        worksheet.set_column('F:F', 12)  # Follow

    writer.close()


def main():
    webtoons = []
    base_url = "https://www.webtoons.com/en/originals?weekday=MONDAY&sortOrder=LIKEIT&webtoonCompleteType=ONGOING"
    html = get_html(base_url)  # Récupérer la valeur de html

    urls = parse_url(html)  # Passer la valeur de html à parse_url

    for url in urls:
        #print(f"Click sur un webtoon : {url}")
        html_webtoon = get_html(url)  # Utiliser la valeur de html pour chaque lien
        webtoons.append(parse_details_page(html_webtoon))
        print(f"Extract Data : {webtoons[-1]}")
        time.sleep(0.1)
    
    df = clean_data(webtoons)
    export_to_excel(df)
    print("Fichier exporté avec succès")

if __name__ == "__main__":
    main()
