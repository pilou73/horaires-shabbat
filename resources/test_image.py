from PIL import Image

try:
    # Ouvre le fichier template.jpg
    img = Image.open("template.jpg")
    print("L'image a été ouverte avec succès !")
    
    # Affiche l'image
    img.show()
except Exception as e:
    print(f"Erreur lors de l'ouverture de l'image : {e}")
