# Made by CrumblyLiquid in 2021

import os
from pathlib import Path
from pptx import Presentation

# Go through every .pptx file in that directory
for filepath in os.listdir(Path(__file__).parent.absolute()):
    if filepath.endswith('.pptx'):
        p = Presentation(filepath)
        text = []
        previous = ""
        skips = ["Zdroje", "Zdroje:", "Obsah", "Obsah:", "Doplňovačka", "Doplňovačka:"]
        for slide in p.slides:
            title = None
            parts = []
            for shape in slide.shapes:
                # Only shapes with text
                if hasattr(shape, "text"):
                    # Check for title shape
                    if (shape == slide.shapes.title and shape.text == slide.shapes.title.text):
                        t = shape.text.replace("\x0b", "")
                        t = t.rstrip(" ").rstrip("\t")
                        if (t != ""):
                            title = t
                    # For every other shape
                    else:
                        for paragraph in shape.text_frame.paragraphs:
                            t = paragraph.text.replace("\x0b", "")
                            t = t.rstrip(" ").rstrip("\t")
                            # Only text with something in it
                            if (t != ""):
                                parts.append("\t- " + t)
            # For first slide only
            if (p.slides.index(slide) == 0 and title is not None):
                text.append(f"----- {title} -----\n")
            # Skip slide
            elif (title in skips):
                continue
            elif (len(parts) > 0):
                if (title is not None):
                    # Add title if it's a different than the last one
                    if (title != previous):
                        text.append(title)
                    # Remove \n if both slides have the same title
                    else:
                        # We remove the last item of the list which should be a ""
                        # thus removing the \n that would othervise be there
                        text.pop(-1)
                    previous = title
                # Add paragraphs
                text.extend(parts)
                # Add slide splitter
                text.append("") # Will turn into \n

        # Save text into a .txt file
        txt_name = filepath.split("/")[-1].replace(".pptx", "") + ".txt"
        txt_path = str(Path(__file__).parent.absolute() / txt_name)
        with open(txt_path, "w", encoding="utf-8") as fp:
            fp.write("\n".join(text))