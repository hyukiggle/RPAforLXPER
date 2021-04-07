from PIL import Image
import pytesseract
from pdf2image import convert_from_path
import os
import os.path


def list_files(path, ext):
    filelist = []
    for name in os.listdir(path):
        if os.path.isfile(os.path.join(path, name)):
            if name.endswith(ext):
                filelist.append(name)
    return filelist

ingredList = list_files('C:/Users/LXPER MINI001/Downloads/모의고사', 'pdf')

for g in range(len(ingredList)):
    name = ingredList[g][:-4]
    os.makedirs('C:/Users/LXPER MINI001/Downloads/pdfTOtext/{}'.format(name))
    os.makedirs('C:/Users/LXPER MINI001/Downloads/pdfTOtext/{}/{}'.format(name,"png"))

    os.chdir('C:/Users/LXPER MINI001/Downloads/pdfTOtext/{}/{}'.format(name, 'png'))
    pages = convert_from_path('C:/Users/LXPER MINI001/Downloads/모의고사/'+ingredList[g], 500, poppler_path = r"C:\Program Files\poppler-21.01.0\Library\bin")
    image_counter = 1
    for page in pages:
        filename = name + "_page" + str(image_counter) + ".png"

        page.save(filename, "PNG")
        image_counter += 1

    filelimit = image_counter - 1
    outfile = "{}.txt".format(name)
    os.chdir('C:/Users/LXPER MINI001/Downloads/pdfTOtext')
    f = open(outfile, "a", encoding='utf-8')
    # Iterate from 1 to total number of pages
    os.chdir('C:/Users/LXPER MINI001/Downloads/pdfTOtext/{}/{}'.format(name, 'png'))
    for i in range(1, filelimit + 1):
        # Set filename to recognize text from
        # Again, these files will be:
        # page_1.jpg
        # page_2.jpg
        # ....
        # page_n.jpg
        filename = name + "_page" + str(i) + ".png"

        # Recognize the text as string in image using pytesserct
        text = str(((pytesseract.image_to_string(Image.open(filename), lang='eng+kor'))))

        # The recognized text is stored in variable text
        # Any string processing may be applied on text
        # The rest of the word is written in the next line
        # Eg: This is a sample text this word here GeeksF-
        # orGeeks is half on first line, remaining on next.
        # To remove this, we replace every '-\n' to ''.
        text = text.replace('-\n', '')

        # Finally, write the processed text to the file.
        f.write(text)

        # Close the file after writing all the text.
    f.close()
    print(str(g+1) + "번째 파일 완료")