import win32com.client
import os

ppt_path = os.path.abspath("presentation/presentation.pptx")
output_dir = os.path.abspath("presentation/slide_images")
os.makedirs(output_dir, exist_ok=True)

app = win32com.client.Dispatch("PowerPoint.Application")
pres = app.Presentations.Open(ppt_path, ReadOnly=True, Untitled=False, WithWindow=False)

try:
    total_slides = pres.Slides.Count
    print(f"Exporting {total_slides} slides to {output_dir}...")
    for i in range(1, total_slides + 1):
        slide = pres.Slides(i)
        out_file = os.path.join(output_dir, f"slide_{i:02d}.png")
        slide.Export(out_file, "PNG", 1920, 1080)
        print(f"Exported: slide_{i:02d}.png")
finally:
    pres.Close()
    app.Quit()

print("All slides exported successfully.")
