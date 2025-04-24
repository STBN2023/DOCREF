import logging


def analyze_placeholders(prs):
    """
    Analyse tous les placeholders dans la présentation.
    Renvoie un dict mapping placeholder nettoyé -> placeholder exact.
    """
    found = {}
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                text = para.text
                idx = 0
                while True:
                    start = text.find('{{', idx)
                    if start == -1:
                        break
                    end = text.find('}}', start)
                    if end == -1:
                        break
                    ph = text[start:end+2]
                    found[ph.strip()] = ph
                    idx = end + 2
    return found


def find_placeholders_in_slide(slide):
    """
    Identique à analyze_placeholders mais pour une seule slide.
    """
    phs = {}
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            text = para.text
            idx = 0
            while True:
                start = text.find('{{', idx)
                if start == -1:
                    break
                end = text.find('}}', start)
                if end == -1:
                    break
                ph = text[start:end+2]
                phs[ph.strip()] = ph
                idx = end + 2
    return phs


def normalize_placeholder(ph):
    """
    Nettoie un placeholder (espaces internes).
    """
    if ph.startswith('{{') and ph.endswith('}}'):
        inner = ph[2:-2].strip()
        return f"{{{{{inner}}}}}"
    return ph


def remplacer_placeholder(slide, placeholder, nouvelle_valeur):
    """
    Remplace un placeholder dans une slide en préservant
    intégralement les styles de police, taille, couleur et alignement.
    Returns True si un remplacement a eu lieu.
    """
    logger = logging.getLogger(__name__)

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            if placeholder not in para.text:
                continue

            # Cas simple : le placeholder est le seul texte du paragraphe
            if para.text.strip() == placeholder.strip():
                # Sauvegarde des styles de chaque run
                runs_styles = []
                for run in para.runs:
                    style = {
                        'font_name': run.font.name,
                        'font_size': run.font.size,
                        'bold': run.font.bold or False,
                        'italic': run.font.italic or False,
                        'underline': run.font.underline or False,
                        'color': run.font.color if run.font.color.type else None,
                        'color_type': run.font.color.type if hasattr(run.font.color, 'type') else None,
                    }
                    runs_styles.append(style)

                # Sauvegarde de l'alignement
                alignment = para.alignment

                # Suppression du contenu existant
                element = para._element
                for r in list(element):
                    element.remove(r)

                # Création d'un nouveau run avec le texte remplacé
                new_run = para.add_run()
                new_run.text = str(nouvelle_valeur)

                # Réapplication du style du premier run
                if runs_styles:
                    style = runs_styles[0]
                    if style['font_name']:
                        new_run.font.name = style['font_name']
                    if style['font_size']:
                        new_run.font.size = style['font_size']
                    new_run.font.bold = style['bold']
                    new_run.font.italic = style['italic']
                    new_run.font.underline = style['underline']
                    if style['color']:
                        try:
                            if style['color_type'] == 1:  # RGB
                                new_run.font.color.rgb = style['color'].rgb
                            elif style['color_type'] == 2:  # Theme
                                new_run.font.color.theme_color = style['color'].theme_color
                                if hasattr(style['color'], 'brightness'):
                                    new_run.font.color.brightness = style['color'].brightness
                            elif style['color_type'] == 3:  # Index
                                new_run.font.color.index = style['color'].index
                        except Exception as e:
                            logger.warning(f"Erreur application couleur: {e}")

                # Réapplication de l'alignement
                para.alignment = alignment
                return True

            # Cas complexe : le placeholder est au sein de plusieurs runs
            # 1. Concaténation du texte et collecte des styles
            full_text = ""
            runs_info = []
            for run in para.runs:
                start_pos = len(full_text)
                full_text += run.text
                end_pos = len(full_text)
                style = {
                    'font_name': run.font.name,
                    'font_size': run.font.size,
                    'bold': run.font.bold or False,
                    'italic': run.font.italic or False,
                    'underline': run.font.underline or False,
                    'color': run.font.color if run.font.color.type else None,
                    'color_type': run.font.color.type if hasattr(run.font.color, 'type') else None,
                }
                runs_info.append({'start': start_pos, 'end': end_pos, 'style': style})

            # 2. Remplacement du texte
            new_text = full_text.replace(placeholder, str(nouvelle_valeur))
            diff = len(new_text) - len(full_text)
            pos0 = full_text.find(placeholder)
            pos1 = pos0 + len(placeholder)

            # 3. Ajustement des positions des runs
            for info in runs_info:
                if info['start'] >= pos1:
                    info['start'] += diff
                    info['end'] += diff
                elif info['start'] < pos1 < info['end']:
                    info['end'] += diff

            # 4. Suppression des runs existants
            element = para._element
            for r in list(element):
                element.remove(r)

            # 5. Reconstruction des runs avec styles
            cursor = 0
            for info in runs_info:
                # Texte intermédiaire
                if info['start'] > cursor:
                    inter_text = new_text[cursor:info['start']]
                    if inter_text:
                        run_mid = para.add_run()
                        run_mid.text = inter_text

                # Texte du run
                run_text = new_text[info['start']:info['end']]
                if run_text:
                    run_new = para.add_run()
                    run_new.text = run_text
                    s = info['style']
                    if s['font_name']:
                        run_new.font.name = s['font_name']
                    if s['font_size']:
                        run_new.font.size = s['font_size']
                    run_new.font.bold = s['bold']
                    run_new.font.italic = s['italic']
                    run_new.font.underline = s['underline']
                    if s['color']:
                        try:
                            if s['color_type'] == 1:
                                run_new.font.color.rgb = s['color'].rgb
                            elif s['color_type'] == 2:
                                run_new.font.color.theme_color = s['color'].theme_color
                                if hasattr(s['color'], 'brightness'):
                                    run_new.font.color.brightness = s['color'].brightness
                            elif s['color_type'] == 3:
                                run_new.font.color.index = s['color'].index
                        except Exception as e:
                            logger.warning(f"Erreur application couleur: {e}")

                cursor = info['end']

            return True

    logger.warning(f"Placeholder '{placeholder}' non trouvé dans cette slide.")
    return False
