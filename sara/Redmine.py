import argparse
import os

from redminelib import Redmine

# url='https://euclid.roe.ac.uk/'
# clef_API='e0302e54d5a17c44591d84483eb66d7d8161275b'
# wiki_doc 3PCF-GC_PAQA_report
# python Redmine.py -k e0302e54d5a17c44591d84483eb66d7d8161275b -w 3PCF-GC_PAQA_report

def set_up_args():
    # Syntaxe de la ligne de commande
    parser= argparse.ArgumentParser()
    parser.add_argument('-k', '--key', metavar='', required=True, help='API key')
    parser.add_argument('-w', '--wiki', metavar='', required=True, help='Document Wikipédia')
    args = parser.parse_args()
    return args

def extract(clef_API,wiki_doc):
    # Authentification
    # Ajouter une vérification ?
    url='https://euclid.roe.ac.uk/'
    redmine = Redmine(url, key=clef_API)
    project = redmine.project.get('quality-tools')
    wiki_page = redmine.wiki_page.get(wiki_doc, project_id=project.identifier)

    # Extraction de la page wiki en html
    wiki_page.export('html', savepath='')

    # Extraction de la page wiki en txt
    # wiki_page.export('txt', savepath='')

    # Conversion de la page wiki en docx
    os.system("pandoc --reference-doc custom-reference.docx %s.html -o %s.docx" % (wiki_doc,wiki_doc))

    # Ouverture du fichier docx
    os.system(".\\%s.docx" % (wiki_doc))


if __name__ == '__main__':
    args=set_up_args()
    extract(args.key,args.wiki)

