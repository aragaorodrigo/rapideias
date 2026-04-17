import zipfile
import xml.etree.ElementTree as ET
import re
import os

def audit():
    file_path = "mentoria_serpro/PremioInovacao/Proposta \u2013 Pr\u00eamio Serpro de Inova\u00e7\u00e3o.odt"
    results = []

    if not os.path.exists(file_path):
        print(f"Erro: Arquivo nao encontrado em {file_path}")
        return

    # 1) Extensao e nome do arquivo
    filename = os.path.basename(file_path)
    ext_ok = filename.endswith(".odt")
    name_ok = "Proposta \u2013 Pr\u00eamio Serpro de Inova\u00e7\u00e3o" in filename
    results.append({
        "item": "1) Extensao e nome",
        "status": "Conforme" if (ext_ok and name_ok) else "Nao conforme",
        "evidencia": f"Nome: {filename}"
    })

    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            content_xml = z.read('content.xml').decode('utf-8')
            styles_xml = z.read('styles.xml').decode('utf-8')
            
            root_content = ET.fromstring(content_xml)
            root_styles = ET.fromstring(styles_xml)
            ns = {
                'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
                'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
                'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
                'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'
            }

            all_text_list = []
            sections = {"Objetivo": "", "Justificativa": "", "Resultados": "", "Palavras-chave": ""}
            current_section = None

            for p in root_content.findall('.//text:p', ns):
                txt = "".join(p.itertext()).strip()
                if not txt: continue
                all_text_list.append(txt)
                
                # Detec\u00e7\u00e3o simples de cabe\u00e7alho de se\u00e7\u00e3o
                found_new_section = False
                for section in sections.keys():
                    if txt.lower().startswith(section.lower()):
                        current_section = section
                        found_new_section = True
                        break
                
                if current_section:
                    sections[current_section] += " " + txt

            # 2) Presen\u00e7a das se\u00e7\u00f5es
            missing = [s for s, content in sections.items() if not content.strip()]
            results.append({
                "item": "2) Presenca secoes",
                "status": "Conforme" if not missing else "Nao conforme",
                "evidencia": f"Faltando: {missing}" if missing else "Todas presentes"
            })

            # 3) Contagem de palavras (Objetivo+Justificativa+Resultados+Palavras-chave)
            total_words = 0
            for s in sections.values():
                total_words += len(re.findall(r'\w+', s))
            
            results.append({
                "item": "3) Palavras total (<=500)",
                "status": "Conforme" if total_words <= 500 else "Nao conforme",
                "evidencia": f"{total_words} palavras"
            })

            # 4) At\u00e9 10 palavras-chave
            kw_text = sections.get("Palavras-chave", "")
            kws = [k for k in re.split(r'[,;.]', kw_text) if k.strip() and k.lower() != "palavras-chave"]
            num_kws = len(kws)
            results.append({
                "item": "4) Ate 10 palavras-chave",
                "status": "Conforme" if num_kws <= 10 else "Nao conforme",
                "evidencia": f"{num_kws} detectadas"
            })

            # 5) Declara\u00e7\u00e3o final
            full_text = " ".join(all_text_list).lower()
            # Procurando por parte da declara\u00e7\u00e3o comum do Serpro
            decl_keys = ["declaro", "autoria", "pl\u00e1gio", "regulamento"]
            found_decl = all(k in full_text for k in ["declaro", "trabalho", "autoria"])
            results.append({
                "item": "5) Declaracao final",
                "status": "Conforme" if found_decl else "Nao conforme",
                "evidencia": "Localizada" if found_decl else "Nao encontrada literalmente"
            })

            # 6) Terceira pessoa e sem ID
            # Regex para 1a pessoa
            pessoal = re.search(r'\b(eu|meu|minha|meus|minhas|n\u00f3s|nosso|nossa|nossos|nossas|elaborei|fiz|fizemos|desenvolvemos)\b', full_text)
            results.append({
                "item": "6) 3a pessoa / Sem ID",
                "status": "Nao conforme" if pessoal else "Conforme",
                "evidencia": f"Termo 1a pessoa: '{pessoal.group()}'" if pessoal else "Ok"
            })

            # 7) Fonte Spranq Eco Sans 12
            # Procurar nos font-face e styles
            font_usage = "Spranq Eco Sans" in styles_xml or "Spranq Eco Sans" in content_xml
            # Procurar por 12pt nos estilos de texto
            size_12 = "12pt" in content_xml or "12pt" in styles_xml
            results.append({
                "item": "7) Fonte Spranq / 12",
                "status": "Conforme" if (font_usage and size_12) else "Verificar",
                "evidencia": f"Spranq: {font_usage}, 12pt: {size_12}"
            })

            # 9) Margens
            page_props = root_styles.findall('.//style:page-layout-properties', ns)
            m_info = "N/A"
            if page_props:
                m_top = page_props[0].get('{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}margin-top')
                m_bot = page_props[0].get('{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}margin-bottom')
                m_lef = page_props[0].get('{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}margin-left')
                m_rig = page_props[0].get('{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}margin-right')
                m_info = f"T:{m_top} B:{m_bot} L:{m_lef} R:{m_rig}"
            
            results.append({
                "item": "9) Margens 3/3/2/2",
                "status": "Verificar",
                "evidencia": m_info
            })

    except Exception as e:
        results.append({"item": "Erro", "status": "Erro", "evidencia": str(e)})

    print("-" * 60)
    print(f"{'ITEM':<25} | {'STATUS':<12} | {'EVIDENCIA'}")
    print("-" * 60)
    for r in results:
        print(f"{r['item']:<25} | {r['status']:<12} | {r['evidencia']}")

audit()
