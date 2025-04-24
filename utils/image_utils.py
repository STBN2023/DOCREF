import os
import logging
from PIL import Image

def creer_image_transparente(chemin='transparent.png', taille=(100,100)):
    """
    Crée une image transparente pour remplacer les blasons inactifs.
    """
    try:
        img = Image.new('RGBA', taille, (255,255,255,0))
        img.save(chemin)
        return chemin
    except Exception as e:
        logging.getLogger(__name__).warning(f"Erreur création image transparente: {e}")
        return None


def remplacer_image(slide, placeholder_name, image_path):
    """
    Remplace un placeholder d'image par un fichier existant.
    """
    logger = logging.getLogger(__name__)
    if not os.path.exists(image_path):
        logger.warning(f"Image non trouvée: {image_path}")
        return False
    target = None
    for shp in slide.shapes:
        if getattr(shp, 'name', '') == placeholder_name:
            target = shp
            break
    if not target:
        for shp in slide.shapes:
            if shp.shape_type == 13:  # picture
                target = shp
                break
    if not target:
        logger.warning(f"Placeholder image '{placeholder_name}' non trouvé.")
        return False
    left, top, width, height = target.left, target.top, target.width, target.height
    el = target._element
    parent = el.getparent()
    parent.remove(el)
    pic = slide.shapes.add_picture(image_path, left, top, width, height)
    try:
        pic.name = placeholder_name
    except:
        pass
    return True