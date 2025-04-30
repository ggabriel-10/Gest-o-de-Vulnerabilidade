import pandas as pd
import requests
import time
from deep_translator import GoogleTranslator

NVD_API_KEY = "25aacf8d-57e5-45c2-8867-28e118b82056"
CISA_EXPLOITED_CVES = set()  # Global para cache da lista da CISA

def load_cisa_exploited_cves_from_json():
    global CISA_EXPLOITED_CVES
    if CISA_EXPLOITED_CVES:
        return CISA_EXPLOITED_CVES  # Usa cache
    url = "https://www.cisa.gov/sites/default/files/feeds/known_exploited_vulnerabilities.json"
    try:
        print("🔄 Baixando catálogo CISA de CVEs exploradas ativamente...")
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        data = response.json()
        CISA_EXPLOITED_CVES = {item['cveID'] for item in data.get('vulnerabilities', []) if 'cveID' in item}
        print(f"✅ {len(CISA_EXPLOITED_CVES)} CVEs carregadas da CISA.")
        return CISA_EXPLOITED_CVES
    except Exception as e:
        print(f"❌ Erro ao baixar o catálogo CISA: {e}")
        return set()

def consulta_nvd(cve_id):
    url = f"https://services.nvd.nist.gov/rest/json/cves/2.0?cveId={cve_id}"
    headers = {
        "apiKey": NVD_API_KEY,
        "User-Agent": "Mozilla/5.0"
    }
    try:
        response = requests.get(url, headers=headers, timeout=25)
        if response.status_code == 429:
            print("⏳ Limite de requisições NVD atingido. Aguardando 30 segundos...")
            time.sleep(30)
            return consulta_nvd(cve_id)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"NVD falhou para {cve_id}: {e}")
        return None

def consulta_circl(cve_id):
    url = f"https://cve.circl.lu/api/cve/{cve_id}"
    try:
        response = requests.get(url, timeout=25)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"CIRCL falhou para {cve_id}: {e}")
        return None

def get_cve_info(cve_id):
    nvd_data = consulta_nvd(cve_id)
    if nvd_data and 'vulnerabilities' in nvd_data:
        vuln = nvd_data['vulnerabilities'][0]['cve']
        description = vuln['descriptions'][0]['value'] if vuln['descriptions'] else ""
        mitigation = vuln.get('cisaRequiredAction', "")
        cvss = vuln.get('metrics', {}).get('cvssMetricV31', [{}])[0].get('cvssData', {}).get('baseScore')
        return {
            'id': cve_id,
            'description': description,
            'cvss': cvss,
            'mitigation': mitigation or "Nenhuma mitigação especificada",
            'source': 'NVD'
        }

    circl_data = consulta_circl(cve_id)
    if circl_data and 'id' in circl_data:
        return {
            'id': circl_data.get('id', cve_id),
            'description': circl_data.get('summary', ""),
            'cvss': circl_data.get('cvss'),
            'mitigation': "Nenhuma mitigação especificada",
            'source': 'CIRCL'
        }

    return None

def safe_translate(translator, text):
    try:
        return translator.translate(text) if text else text
    except Exception as e:
        print(f"Erro ao traduzir '{text}': {e}")
        return text

def classificar_criticidade(cvss, explorado):
    if explorado:
        return "Crítica"
    elif cvss is not None:
        if cvss >= 8:
            return "Alta"
        elif cvss >= 5:
            return "Média"
        else:
            return "Baixa"
    return "Desconhecida"

def main():
    try:
        df = pd.read_excel('base.xlsx', header=None, names=['CVE'])
    except FileNotFoundError:
        print("❌ Erro: O arquivo 'base.xlsx' não foi encontrado.")
        return

    cisa_exploited = load_cisa_exploited_cves_from_json()
    cve_ids = df['CVE'].dropna().astype(str).tolist()
    translator = GoogleTranslator(source='en', target='pt')
    data = []

    for cve_id in cve_ids:
        cve_info = get_cve_info(cve_id)

        if not cve_info:
            print(f"{cve_id} não encontrado em nenhuma API.")
            continue

        desc = safe_translate(translator, cve_info['description'])

        # Traduz mitigação somente se houver algo útil
        if cve_info['mitigation'] and "Nenhuma mitigação especificada" not in cve_info['mitigation']:
            mitig = safe_translate(translator, cve_info['mitigation'])
        else:
            mitig = "Nenhuma mitigação especificada"

        explorado = cve_id in cisa_exploited
        criticidade = classificar_criticidade(cve_info['cvss'], explorado)

        print(f"{cve_id} - {criticidade} ({'Explorado' if explorado else 'Não explorado'}) via {cve_info['source']}")

        data.append({
            'CVE': cve_id,
            'Descrição': desc,
            'CVSS': cve_info['cvss'],
            'Mitigação': mitig,
            'Exploração Ativa': "Sim" if explorado else "Não",
            'Criticidade Real': criticidade,
            'Fonte': cve_info['source']
        })

    if data:
        df_out = pd.DataFrame(data)
        df_out.to_csv('output.csv', index=False, encoding='utf-8-sig')
        print("✅ Resultados salvos em 'output.csv'.")
    else:
        print("⚠️ Nenhum dado válido extraído.")

if __name__ == "__main__":
    main()
