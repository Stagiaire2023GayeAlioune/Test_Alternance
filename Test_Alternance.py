from pptx import Presentation
import csv
import unittest
import os

def extract_table_data(pptx_file):
    # Charger le fichier PowerPoint
    pptx_file = "C:/Users/math/Desktop/mes_cv_et_lettre/Exemple FASEP.pptx"
    power = Presentation(pptx_file)

    # Ouvrir un fichier en mode écriture pour écrire les données extraites au format CSV
    with open("File_To_Extract_ppt.csv", "w", newline='', encoding='utf-8') as csv_file:
        # Initialiser un writer pour écrire dans le fichier CSV
        csv_writer = csv.writer(csv_file)
        
        # Parcourir toutes les diapositives du fichier PowerPoint
        for slide in power.slides:
            # Parcourir toutes les formes sur la diapositive
            for shape in slide.shapes:
                # Vérifier si la forme est un tableau
                if shape.has_table:
                    # Accéder au tableau
                    tableau = shape.table
                    # Parcourir chaque ligne du tableau
                    for row in tableau.rows:
                        # Initialiser une variable pour vérifier si une des phrases est trouvée dans la ligne
                        phrase_trouvée = False
                        # Parcourir chaque cellule de la ligne
                        for cell in row.cells:
                            # Accéder au texte de la cellule
                            if hasattr(cell, "text"):
                                texte_cellule = cell.text.strip() if cell.text else ""
                            else:
                                # Si la cellule est un GraphicFrame, accéder à son contenu
                                texte_cellule = ""
                                for paragraph in cell.text_frame.paragraphs:
                                    texte_cellule += paragraph.text.strip() + " "
                            # Vérifier si l'une des phrases recherchées est présente dans le texte de la cellule
                            if "Date de signature de la convention Natixis" in texte_cellule or \
                               "Montant et date de paiement de l'acompte" in texte_cellule or \
                               "Avis sur le versement intermédiaire" in texte_cellule or \
                               "Date de l'avis" in texte_cellule:
                                phrase_trouvée = True
                                break  # Sortir de la boucle si l'une des phrases est trouvée dans cette ligne
                        
                        # Si l'une des phrases est trouvée dans la ligne, récupérer la ligne entière
                        if phrase_trouvée:
                            # Accéder au texte de chaque cellule dans la ligne et les ajouter à une liste
                            ligne_tableau = [cell.text.strip() if cell.text else "" for cell in row.cells]
                            # Écrire la ligne du tableau dans le fichier CSV
                            csv_writer.writerow(ligne_tableau)

class TestExtractTableData(unittest.TestCase):
    def test_extract_table_data(self):
        # Testez avec un fichier PowerPoint valide
        extract_table_data("C:/Users/math/Desktop/mes_cv_et_lettre/Exemple FASEP.pptx")
        # Assurez-vous que le fichier de sortie existe
        self.assertTrue(os.path.exists("File_To_Extract_ppt.csv"))
        extract_table_data("fichier_inexistant.pptx")
        # Assurez-vous que le fichier de sortie n'existe pas
        self.assertFalse(os.path.exists("File_To_Extract_ppt.csv"))
if __name__ == '__main__':
    unittest.main()