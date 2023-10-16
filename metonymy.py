import nltk
from nltk.corpus import wordnet as wn
import pandas as pd

nltk.download('wordnet')

def identify_metonymy(word):
    synsets = wn.synsets(word)
    polysemous_synsets = [synset for synset in synsets if len(synset.lemmas()) > 1]

    metonymies = []
    for synset in polysemous_synsets:
        lemma = synset.lemmas()[0].name()

        cross_reference_relations = []
        for other_synset in polysemous_synsets:
            if other_synset != synset:
                cross_reference_relations.append((synset.name(), other_synset.name()))

        for relation in cross_reference_relations:
            relation_synset = wn.synset(relation[0])
            relation_hypernyms = relation_synset.hypernyms()

            for hypernym in relation_hypernyms:
                metonymies.append((lemma, hypernym.lemmas()[0].name()))

    return metonymies

file_path = r"C:\Users\de'l'l\Desktop\testzhuan\sentence-LloydList.txt"
output_path = r"C:\Users\de'l'l\Desktop\testzhuan\metonymies.xlsx"

with open(file_path, 'r', encoding='utf-8') as file:
    text = file.read()

words = text.split()
found_metonymies = []
for word in words:
    metonymies = identify_metonymy(word)
    found_metonymies.extend(metonymies)

if len(found_metonymies) > 0:
    max_rows_per_sheet = 1048576
    num_sheets = -(-len(found_metonymies) // max_rows_per_sheet)  # 向上取整获取所需的工作表数量

    with pd.ExcelWriter(output_path) as writer:
        for i in range(num_sheets):
            start_idx = i * max_rows_per_sheet
            end_idx = start_idx + max_rows_per_sheet
            chunk = found_metonymies[start_idx:end_idx]

            data = {
                'Word': [metonymy[0] for metonymy in chunk],
                'Metonymy': [metonymy[1] for metonymy in chunk]
            }
            df = pd.DataFrame(data)
            sheet_name = f'Sheet{i+1}'
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"The metonymies have been saved to {output_path} with {num_sheets} sheet(s).")
else:
    print("The text does not contain any metonymies.")
