import pypdfium2 as pdfium
import sys, os

files = [
    ("83", r"data/Заявка №83 от 24.09.2026г Жилой дом Шатура.pdf"),
    ("53", r"data/СЗ 53 от 24.09.2026г Жилые дома пос. Климовск Южный.pdf"),
    ("44", r"data/Заявка №44 от 23.09.2026г Видное ДОУ.pdf"),
]

out_dir = "tmp/zayavki_render"
os.makedirs(out_dir, exist_ok=True)

for tag, path in files:
    pdf = pdfium.PdfDocument(path)
    n = len(pdf)
    print(tag, path, "pages:", n)
    for i in range(n):
        page = pdf[i]
        bitmap = page.render(scale=2.8)
        pil_image = bitmap.to_pil()
        out_path = os.path.join(out_dir, f"{tag}_p{i+1}.png")
        pil_image.save(out_path)
        print("saved", out_path)
