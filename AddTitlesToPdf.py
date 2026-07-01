from pypdf import PdfReader, PdfWriter
from pypdf.generic import Fit
import sys


reader = PdfReader(sys.argv[1])
writer = PdfWriter()
writer.append(reader)


metadata_list = dict(reader.metadata)
"""
A dictionary that contains all the metadata of the source PDF. Add new keys (with `/` before their name) to the dictionary to add new tags.
"""
metadata_list["/Creator"] = "Microsoft Word"

settings = {
    "write_metadata": True, # Add author, title and creator tags to the output file
    "top_pt": 2, # When creating the link to the bookmark, leave *this* points above the title
    "left_pt": 2, # When creating the link to the bookmark, leave *this* points at the left of the title
    "put_wrong_outline_as_children_of_previous_item": True, # If the outline level (the one of the heading) is impossible (ex: from outline 2 we go to outline 7) and this option is true, it'll be added as a child of the outline that has been previously added.
    "from_extract": None, # If passed, an array of the pages that the output PDF file includes of the Word file. This is used so that the script can handle extracts of some pages.
    "extract_count_starts_from": 0, # All the entries of the `from_extract` array will be increased of this number.
    "output_name": sys.argv[1] # The name of the output file
}

outline_items = {}
"""
A dictionary that contains as a key the outline level of the string, and as a value the last `outline_item` oject that was added at that position
"""

for i in range(3, len(sys.argv)): 
    match(sys.argv[i]):
        case "--skip-metadata":
            settings["write_metadata"] = False
        case "--top-points":
            settings["top_pt"] = int(sys.argv[i+1])
        case "--left-points":
            settings["left_pt"] = int(sys.argv[i+1])
        case "--put-wrong-outline-first":
            settings["put_wrong_outline_as_children_of_previous_item"] = False
        case "--extract":
            settings["from_extract"] = list()
            # Supported syntax: "a-b,c"
            for num in sys.argv[i+1].split(","):
                dash = num.split("-")
                if len(dash) == 1:
                    settings["from_extract"].append(int(dash[0].strip()) + settings["extract_count_starts_from"])
                else:
                    settings["from_extract"].extend(range(int(dash[0].strip()) + settings["extract_count_starts_from"], int(dash[1].strip()) + settings["extract_count_starts_from"]))
        case "--start-extract-count-from":
            settings["extract_count_starts_from"] = int(sys.argv[i+1])
            if settings["from_extract"] != None: # We need to increase the numbers that have already been added to the from_extract array
                settings["from_extract"] = list(map(lambda x: x + settings["extract_count_starts_from"], settings["from_extract"]))
        case "--output-name":
            settings["output_name"] = sys.argv[i+1]

with open(sys.argv[2], 'r', encoding="mac_roman") as file:
    # Let's keep track of the current line so that we can add metadata to the file
    line_position = -1
    prev_parent = -1
    line = file.readline()
    while line:
        line_position = line_position + 1
        if line_position == 0 and line.strip() != "": # First line = title name
            metadata_list["/Title"] = line
        elif line_position == 1 and line.strip() != "": # Second line = author name
            metadata_list["/Author"] = line
        else: # Add outline item
            info = line.split(";")
            if len(info) > 4: # Syntax: Outline level; Page; Horizontal position; Vertical position; Text
                text = ";".join(info[4:])
                if text.strip() != "" and (settings["from_extract"] == None or int(info[1]) in settings["from_extract"]): # if settings["from_extract"] == None, all the pages are available in the output PDF file. Otherwise, we need to check that the current page has been added to the extract.
                    page_num = settings["from_extract"].index(int(info[1])) if settings["from_extract"] != None else (int(info[1]) - 1) # Page number. If "from_extract" is passed, we'll get the actual page position in the output file.
                    y_pdf_coords = float(reader.pages[page_num].mediabox.height) - float(info[3].replace(",", ".")) + settings["top_pt"] # The y coordinate of PDF files starts from the bottom, and not from the top
                    correct_outline_level = prev_parent >= (int(info[0]) - 1) # Checks that the file structure is valid (eg. if the previous item had an outline of 2, the new item cannot have an outline of 7; but if the previous item had an outline of 7, the new item can have an outline of 2)
                    outline_items[info[0]] = writer.add_outline_item(
                        text, 
                        page_num,
                        fit=Fit.xyz(left=(float(info[2].replace(",", ".")) - settings["left_pt"]), top=y_pdf_coords), # Position of the outline
                        parent=outline_items[str(int(info[0]) - 1)] if correct_outline_level and str(int(info[0]) - 1) in outline_items else outline_items[str(prev_parent)] if not correct_outline_level and str(prev_parent) in outline_items and settings["put_wrong_outline_as_children_of_previous_item"] else  None # If possible, add it in a nested list. We check first that the structure is valid, and then that a parent can actually be found. If it's not valid, the script might use the last added outline as a parent.
                    )
                    prev_parent = int(info[0])
        line = file.readline()

if settings["write_metadata"]: writer.add_metadata(metadata_list)

with open(settings["output_name"], "wb") as output:
    writer.write(output)