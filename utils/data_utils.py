import pandas as pd
import numpy as np

def convertir_en_numerique(valeur):
    """
    Convertit une valeur en type numérique si possible,
    mais ignore les valeurs textuelles qui ne sont pas clairement numériques.
    """
    try:
        if isinstance(valeur, str) and (not valeur.replace('.', '', 1).replace(',', '', 1).isdigit()):
            return valeur
        return pd.to_numeric(valeur)
    except (ValueError, TypeError):
        return valeur
    except Exception:
        return valeur


def formater_valeur(valeur, type_format=None):
    """
    Formate une valeur selon un type (monétaire, surface, date...).
    """
    if pd.isna(valeur):
        return ''
    # On étend la reconnaissance aux types numpy.integer et numpy.floating
    if type_format == 'monétaire' and isinstance(valeur, (int, float, np.integer, np.floating)):
        # Exemple : 100000 -> "100 000,00 €"
        return f"{valeur:,.2f} €".replace(',', ' ').replace('.', ',')
    if type_format == 'surface' and isinstance(valeur, (int, float, np.integer, np.floating)):
        return f"{valeur:,} m²".replace(',', ' ')
    if isinstance(valeur, (pd.Timestamp, np.datetime64)):
        return pd.to_datetime(valeur).strftime('%d/%m/%Y')
    return str(valeur)