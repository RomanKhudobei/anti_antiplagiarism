import sys
import os
import shutil
import distutils.dir_util
import distutils.file_util
import re
import random
import copy
import zipfile

import lxml.etree as ET
import pymsgbox as pmb


def add_content_types(target):
    tree = ET.parse(os.path.join(target, '[Content_Types].xml'), parser=ET.XMLParser(remove_blank_text=True))
    root = tree.getroot()

    insertion = ET.fromstring('<Default ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont" Extension="odttf"/>')
    root.append(insertion)

    tree.write(os.path.join(target, '[Content_Types].xml'), xml_declaration=True, encoding='utf-8', standalone=True)


def copy_custom_xml(src, target):
    distutils.dir_util.copy_tree(
        os.path.join(src, 'customXml'),
        os.path.join(target, 'customXml')
    )


def inject_font(src, target):
    copy_font(src, target)
    extend_font_table(target)


def copy_font(src, target):
    distutils.dir_util.copy_tree(
        os.path.join(src, 'word/fonts'),
        os.path.join(target, 'word/fonts')
    )

    shutil.copy(
        os.path.join(src, 'word/_rels/fontTable.xml.rels'),
        os.path.join(target, 'word/_rels/fontTable.xml.rels')
    )


def extend_font_table(target):
    tree = ET.parse(os.path.join(target, 'word/fontTable.xml'))
    root = tree.getroot()

    insertion = ET.fromstring("""
        <w:fonts mc:Ignorable="w14 w15 w16se" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex">
            <w:font w:name="Tirnes New Roman">
                <w:charset w:val="CC"/>
                <w:family w:val="roman"/>
                <w:pitch w:val="variable"/>
                <w:sig w:csb0="000001FF" w:csb1="00000000" w:usb0="E0002AFF" w:usb1="C0007841" w:usb2="00000009" w:usb3="00000000"/>
                <w:embedRegular r:id="rId1" w:fontKey="{312D396F-96FB-4265-A766-1D2F18E175CC}"/>
                <w:embedBold r:id="rId2" w:fontKey="{819CC672-0EB7-4F73-A5AD-89DA36D51272}"/>
                <w:embedItalic r:id="rId3" w:fontKey="{6910D700-D703-438C-80BD-1451DA89D10A}"/>
            </w:font>
        </w:fonts>
    """, parser=ET.XMLParser(remove_blank_text=True)).find(f"{{{root.nsmap.get('w')}}}font")

    root.append(insertion)

    tree.write(os.path.join(target, 'word/fontTable.xml'), xml_declaration=True, encoding='utf-8', standalone=True)


def change_file_settings(target):
    tree = ET.parse(os.path.join(target, 'word/settings.xml'))
    root = tree.getroot()

    insertion = ET.fromstring("""
        <w:settings mc:Ignorable="w14 w15 w16se" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex">
            <w:embedTrueTypeFonts/>
            <w:saveSubsetFonts/>
            <w:hideGrammaticalErrors/>
            <w:hideSpellingErrors/>
            <w:noPunctuationKerning/>
            <w:doNotIncludeSubdocsInStats/>
        </w:settings>
    """, parser=ET.XMLParser(remove_blank_text=True))

    root.extend(insertion.getchildren())

    tree.write(os.path.join(target, 'word/settings.xml'), xml_declaration=True, encoding='utf-8', standalone=True)


def make_unique(target, percentage):
    # total_words = get_total_words(target)
    # words_left = (percentage / 100) * total_words

    tree = ET.parse(os.path.join(target, 'word/document.xml'), parser=ET.XMLParser(remove_blank_text=True))
    root = tree.getroot()

    for origin_run in root.iter(tag=f"{{{root.nsmap.get('w')}}}r"):

        # if words_left <= 0:
        #     break

        text = origin_run.find(f"{{{origin_run.nsmap.get('w')}}}t")

        if text is None:
            continue

        if not getattr(text, 'text', None):
            text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            text.text = ' '
            continue

        text = text.text

        words_count = len(re.findall(r'\w+', text))

        if words_count <= 1 and len(text) <= 5:
            continue

        splitted_elements = []
        revision_id = get_revision_id(origin_run)  # rsidRPr attribute
        run_properties = get_run_properties(origin_run)    # <w:rPr> tag

        for index, word in enumerate(re.split(r'(\s+)', text)):

            space_probability = random.randint(0, 100)

            if space_probability > percentage:
                run = ET.Element(f"{{{origin_run.nsmap.get('w')}}}r", nsmap={'w': origin_run.nsmap.get('w')})

                if revision_id is not None:
                    run.set(f"{{{origin_run.nsmap.get('w')}}}rsidRPr", revision_id)

                if run_properties is not None:
                    run.append(copy.deepcopy(run_properties))

                text = ET.Element(f"{{{origin_run.nsmap.get('w')}}}t")
                text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                text.text = word

                run.append(text)
                splitted_elements.append(run)
                continue

            is_first_word = (index == 0)

            if not re.findall(r'\w+', word) or len(word) <= 3:

                if is_first_word:
                    run = copy.deepcopy(origin_run)
                    text = run.find(f"{{{origin_run.nsmap.get('w')}}}t")
                    text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    text.text = word
                else:
                    run = ET.Element(f"{{{origin_run.nsmap.get('w')}}}r", nsmap={'w': origin_run.nsmap.get('w')})

                    if revision_id is not None:
                        run.set(f"{{{origin_run.nsmap.get('w')}}}rsidRPr", revision_id)

                    if run_properties is not None:
                        run.append(copy.deepcopy(run_properties))

                    text = ET.Element(f"{{{origin_run.nsmap.get('w')}}}t")
                    text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    text.text = word

                    run.append(text)

                splitted_elements.append(run)
                continue

            split = random.randint(1, len(word)-1)
            word1, word2 = word[:split], word[split:]

            if is_first_word:
                run1 = copy.deepcopy(origin_run)
                text1 = run1.find(f"{{{origin_run.nsmap.get('w')}}}t")
                text1.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                text1.text = word1
            else:
                run1 = ET.Element(f"{{{origin_run.nsmap.get('w')}}}r", nsmap={'w': origin_run.nsmap.get('w')})

                if revision_id is not None:
                    run1.set(f"{{{origin_run.nsmap.get('w')}}}rsidRPr", revision_id)

                if run_properties is not None:
                    run1.append(copy.deepcopy(run_properties))

                text1 = ET.Element(f"{{{origin_run.nsmap.get('w')}}}t")
                text1.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                text1.text = word1

                run1.append(text1)

            # form up space
            empty_space = ET.Element(f"{{{origin_run.nsmap.get('w')}}}r", nsmap={'w': origin_run.nsmap.get('w')})

            if revision_id is not None:
                empty_space.set(f"{{{origin_run.nsmap.get('w')}}}rsidRPr", revision_id)

            if run_properties is not None:
                empty_space_run_properties = copy.deepcopy(run_properties)
            else:
                empty_space_run_properties = ET.Element(f"{{{origin_run.nsmap.get('w')}}}rPr", nsmap={'w': origin_run.nsmap.get('w')})

            font_tag = ET.Element(f"{{{origin_run.nsmap.get('w')}}}rFonts")
            font_tag.set(f"{{{origin_run.nsmap.get('w')}}}ascii", 'Tirnes New Roman')
            font_tag.set(f"{{{origin_run.nsmap.get('w')}}}hAnsi", 'Tirnes New Roman')

            empty_space_run_properties.append(font_tag)

            no_proof_tag = ET.Element(f"{{{origin_run.nsmap.get('w')}}}noProof")

            empty_space_run_properties.append(no_proof_tag)

            empty_space.append(empty_space_run_properties)

            empty_space_text_tag = ET.Element(f"{{{origin_run.nsmap.get('w')}}}t")
            empty_space_text_tag.text = ' '     # ATTENTION-ATTENTION: THIS IS UNBREAKABLE SPACE (ctrl+shift+space in MS Word)

            empty_space.append(empty_space_text_tag)

            # form up word2
            run2 = ET.Element(f"{{{origin_run.nsmap.get('w')}}}r", nsmap={'w': origin_run.nsmap.get('w')})

            if revision_id is not None:
                run2.set(f"{{{origin_run.nsmap.get('w')}}}rsidRPr", revision_id)

            if run_properties is not None:
                run2.append(copy.deepcopy(run_properties))

            text2 = ET.Element(f"{{{origin_run.nsmap.get('w')}}}t")
            text2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            text2.text = word2

            run2.append(text2)

            # append everything to splitted elements
            splitted_elements.extend([run1, empty_space, run2])

        for e in reversed(splitted_elements):
            origin_run.addnext(e)

        origin_run.getparent().remove(origin_run)
        # words_left -= words_count

    tree.write(os.path.join(target, 'word/document.xml'), xml_declaration=True, encoding='utf-8', standalone=True)


# def get_total_words(target):
#     tree = ET.parse(os.path.join(target, 'docProps/app.xml'), parser=ET.XMLParser(encoding='utf-8', remove_blank_text=True))
#     root = tree.getroot()

#     total_words = root.find(f'{{{root.nsmap.get(None)}}}Words')

#     total_words = getattr(total_words, 'text', None)
#     return total_words and int(total_words)


def get_revision_id(run):
    revision_attr = 'rsidRPr'

    for key in run.keys():
        if revision_attr in key:
            revision_attr = key
            break

    return run.get(revision_attr)


def get_run_properties(run):
    return run.find(f"{{{run.nsmap.get('w')}}}rPr")


def main():
    origin_dir_path = pmb.prompt(
        title='Задання початкових параметрів',
        text='Вкажіть шлях до папки з файлом: (УВАГА! Для вставки через "Ctrl+V" необхідна англійська розкладка клавіатури)'
    )

    filename = pmb.prompt(
        title='Задання початкових параметрів',
        text='Вкажіть назву файлу: (Файл повинен мати розширення ".docx")'
    )

    percentage = pmb.prompt(
        title='Задання початкових параметрів',
        text='Вкажіть ймовірність вставки пробілу в слово: (Від 0 до 100 без знаку "%")'
    )

    try:
        percentage = int(percentage)
    except (ValueError, TypeError):
        pmb.alert(
            title='Помилка',
            text=f'Невірно вказано процент унікальності "{percentage}"'
        )
        sys.exit(0)

    if 100 < percentage < 0:
        pmb.alert(
            title='Помилка',
            text=f'Невірно вказано процент унікальності "{percentage}" (Приймаються значення від 0 до 100, без знаку "%")'
        )
        sys.exit(0)

    if not filename.endswith('.docx'):
        filename += '.docx'

    origin_doc_path = os.path.join(origin_dir_path, filename)

    if not os.path.exists(origin_doc_path):
        pmb.alert(
            title='Помилка',
            text='Вказаного файлу не знайдено або файл має невірне розширення.' +
            f'Пошук здійснювався за наступним шляхом: "{origin_doc_path}"'
        )

    # Copy to operate not on original file
    target_doc_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), filename)
    shutil.copy(
        origin_doc_path,
        target_doc_path
    )

    target_file_path, extension = os.path.splitext(target_doc_path)    
    target_zip_path = target_file_path + '.zip'
    os.rename(target_doc_path, target_zip_path)

    with zipfile.ZipFile(target_zip_path, 'r') as zip_ref:
        target_dir_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'target')
        zip_ref.extractall(target_dir_path)
    
    os.remove(target_zip_path)

    try:
        base_path = sys._MEIPASS
    except:
        base_path = os.path.abspath(".")

    source_dir_path = os.path.join(base_path, 'source')

    add_content_types(target_dir_path)
    copy_custom_xml(source_dir_path, target_dir_path)

    inject_font(source_dir_path, target_dir_path)

    change_file_settings(target_dir_path)

    make_unique(target_dir_path, percentage)

    output_file_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), os.path.basename(target_file_path) + ' [УНІКАЛЬНИЙ]')
    shutil.make_archive(output_file_path, 'zip', target_dir_path)

    shutil.move(output_file_path + '.zip', output_file_path + '.docx')

    shutil.rmtree(target_dir_path)

    pmb.alert(
        title='Успіх',
        text=f'Вказаний файл успішно оброблено. Збережено за шляхом: "{output_file_path}"'
    )


if __name__ == '__main__':
    main()
