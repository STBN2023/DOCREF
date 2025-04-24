import os
import logging

def gerer_blasons_ameliore(slide, projet, mapping_blasons, logos_dir, valeur_active='x'):
    """
    Affiche dynamiquement les blasons actifs autour d'un centre.
    """
    logger = logging.getLogger(__name__)
    actifs = []
    for col, filename in mapping_blasons.items():
        if str(projet.get(col, '')).strip().lower() == valeur_active.lower():
            path = os.path.join(logos_dir, filename)
            if os.path.exists(path):
                actifs.append(path)
    if not actifs:
        return []
    # déterminer centre
    centre = None
    for shp in slide.shapes:
        if getattr(shp, 'name', '') == 'blason_centre':
            centre = shp
            break
    if centre:
        cx = centre.left + centre.width//2
        cy = centre.top + centre.height//2
        try:
            centre._element.getparent().remove(centre._element)
        except:
            pass
        w, h = centre.width, centre.height
    else:
        cx = slide.slide_width//2
        cy = slide.slide_height//2
        w, h = 1000000, 1000000
    # positionner en ligne verticale
    esp = h * 1.2
    for i, path in enumerate(actifs):
        x = cx - w//2
        y = cy - (len(actifs)//2)*esp + i*esp - h//2
        slide.shapes.add_picture(path, x, y, w, h)
    return actifs