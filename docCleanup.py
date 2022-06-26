import math

##################### CLI ######################
import sys

class Arguments:
    old_file = "old"
    new_file = "new"
    factor = "factor"

def print_error(value: str):
    sys.exit(value)


kw_dict = {}
for arg in sys.argv[1:]:
    if '=' in arg:
        sep = arg.find('=')
        key, value = arg[:sep], arg[sep + 1:]
        kw_dict[key] = value

if len(kw_dict) == 0:
    print_error("Missing Arguments. Example: ([] means optional)\npython docCleanup.py old=OldDoc.docx new=NewDoc.docx [factor=.3]")


start_doc = print_error("include 'old=docName.docx'") if Arguments.old_file not in kw_dict else kw_dict.get(Arguments.old_file)
end_doc = print_error("include 'new=docName.docx'") if Arguments.new_file not in kw_dict else kw_dict.get(Arguments.new_file)
factor = .3 if Arguments.factor not in kw_dict else float(kw_dict.get(Arguments.factor))

#################### END CLI #####################

# modify document
from docx import Document
document = Document(start_doc)

# shrink photos
print ("Shrinking photos...")
for shape in document.inline_shapes:

    shape.height = math.trunc(float(shape.height * factor)) 
    shape.width = math.trunc(float(shape.width * factor))
print("Done shrinking photos!")

# remove whitespace
print("Removing extra whitespace...")
import re
for p in document.paragraphs:
    if '/n/n' in p.text and re.search('[a-zA-Z]', p.text):
        p.text = re.sub(r'\n\n+', '\n\n', p.text)
print("Done removing extra whitespace!")


# complete document
print("Saving your document...")
document.save(end_doc)
print("All done!")
