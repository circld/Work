import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import re

# words to strip out
noisy_words = ['invest', 'asset', 'manager', 'partners', 'limited', 'capital'
                , 'corp', 'associate', 'research', 'financial', 'bank', 'management'
                , 'fund', 'group', 'plc', 'ltd', 'company', 'holding', 'service'
                , 'portfolio', 'banco', 'gruppe', 'investment', 'finance'
              ]
# helper functions
def match_lists(list1, list2, precision=90):
    '''
    Takes two lists
    Returns dict of matches (if best match scores above precision)
        and non-matches (otherwise); both dicts have key = list1
        element and value = best match
    '''
    in_both = {}
    not_in_both = {}
    for manager1 in list1:
        best_match = process.extractBests(manager1, list2, limit=1)
        if best_match[0][1] > precision:
            in_both[manager1] = best_match[0][0]
        else:
            not_in_both[manager1] = best_match[0][0]
    return [in_both, not_in_both]

def get_match(word, word_list, precision=90):
    if len(word) == 0:
        return ''
    best_match = process.extractBests(word, word_list, limit=1)
    if best_match[0][1] <= precision:
        best_match = ''
    return best_match

def remove_noisy_words(word):
    for stem in noisy_words:
        pattern = re.compile(stem + 's?', flags=re.IGNORECASE)
        word = re.sub(pattern, '', word)
    return word

# import data
ff_raw = pd.read_excel('./Desktop/managers.xlsx', 'FF')
si_raw = pd.read_excel('./Desktop/managers.xlsx', 'SI')
si_total_raw = pd.read_excel('./Desktop/managers.xlsx', 'si_total')

# drop NA values
si = si_raw.dropna()
ff = ff_raw.dropna()
si_total = si_total_raw.dropna()

# clean and standardize formatting using fuzz module
ff['CleanManagerName'] = ff['ManagerName'].map(fuzz.full_process)
si['CleanManagerName'] = si['ManagerName'].map(fuzz.full_process)
si_total['CleanManagerName'] = si_total['ManagerName'].map(fuzz.full_process)

# remove common words with little signalling value
ff['ManagerStem'] = ff['CleanManagerName'].map(remove_noisy_words)
si['ManagerStem'] = si['CleanManagerName'].map(remove_noisy_words)
si_total['ManagerStem'] = si_total['ManagerName'].map(remove_noisy_words)

# create lists
si_list = [i for i in si['ManagerStem']]
ff_list = [i for i in ff['ManagerStem']]
si_total_list = [i for i in si_total['ManagerStem']]


if __name__ == '__main__':
    # creating column with best match (according to precision)
    ff['ManagerSI'] = ff.ManagerStem.map(lambda x : get_match(x, si_list))
    # unmatched FF CB managers
    ff_only = ff[ff['ManagerSI'] == '']
    ff_only['LocalManagerSI'] = ff_only.ManagerStem.map(lambda x : get_match(x, si_total_list))
    ff_only.to_csv('Desktop/ff_only.csv', encoding='utf-8')
    ff.to_csv('Desktop/ff_analysis.csv', encoding='utf-8')
    #ff_both, ff_only = match_lists(ff_list, si_list, 80)
    #si_both, si_only = match_lists(si_list, ff_list)
    #ff_total_both, ff_total_only = match_lists(ff_only.keys(), si_total_list)
    
