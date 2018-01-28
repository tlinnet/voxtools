# http://scikit-learn.org/stable/tutorial/text_analytics/working_with_text_data.html
# https://towardsdatascience.com/machine-learning-nlp-text-classification-using-scikit-learn-python-and-nltk-c52b92a7c73a

# Load dataset
from sklearn.datasets import fetch_20newsgroups
import random
from openpyxl import Workbook

# Dataset: http://qwone.com/~jason/20Newsgroups/
cat_comp = ['comp.graphics','comp.os.ms-windows.misc','comp.sys.ibm.pc.hardware','comp.sys.mac.hardware','comp.windows.x']
cat_rec = ['rec.autos','rec.motorcycles','rec.sport.baseball','rec.sport.hockey']
cat_sci = ['sci.crypt','sci.electronics','sci.med','sci.space']
cat_misc = ['misc.forsale']
cat_talk = ['talk.politics.misc','talk.politics.guns','talk.politics.mideast','talk.religion.misc']
cat_alt = ['alt.atheism']
cat_soc = ['soc.religion.christian']
cat_all = cat_comp + cat_rec + cat_sci + cat_misc + cat_talk + cat_alt + cat_soc
print("Number of categories is: %i"%len(cat_all))
number_of_samples = 4
categories_ran = random.choices(population=cat_all, k=number_of_samples)
print(categories_ran)

# Load the training data. We will load the test data later
categories = ['alt.atheism', 'soc.religion.christian','comp.graphics', 'sci.med']
twenty_train = fetch_20newsgroups(subset='train', categories=categories, shuffle=True, random_state=42)

print("Number of sentences is: %i\n"%len(twenty_train.data))
print("First sentence is:")
print("########################################################################")
print(twenty_train.data[0])
print("------------------------------------------------------------------------")
print("First sentence is targeted to:")
print(twenty_train.target_names[twenty_train.target[0]])
print("########################################################################")
print()

# Create workbooks
wb = Workbook()
ws = wb.active
# Write header
ws.append(['sentences', 'train', 'classify', 'test', 'correct_train', 'correct_test'])
N_train = 50
N_train_i = 0
N_test = 20
N_test_i = 0
N_tot = 300
print("Creating %i sentences, where %i is training and %i is for test"%(N_tot, N_train, N_test))
print("------------------------------------------------------------------------")
print()

# Loop over total
for i in range(N_tot):
    sent = twenty_train.data[i]
    # encode
    #sent = sent.encode('ascii', 'ignore').decode("utf-8")
    sent = sent.encode('unicode_escape').decode('utf-8')
    cl = twenty_train.target_names[twenty_train.target[i]]
    if N_train_i < N_train:
        row = [sent, cl, None, None]
        N_train_i += 1
    elif N_test_i < N_test:
        row = [sent, None, None, cl]
        N_test_i += 1
    else:
        row = [sent]
    # Append row
    ws.append(row)

wb_fname = 'kodning02.xlsx'
wb.save(wb_fname)
print("Workbook saved!\n")
print(wb_fname)
    