'''
This script, when run from the console, takes a CSV of data (including
FundName and ISIN columns) and returns a CSV with ISIN and minimum fuzzy string
match score.

Takes two command line arguments:
    location of first CSV
    directory to save output CSV

Example Usage:
    $ python MultiNameISIN.py /c/Users/YourName/file1.csv /c/
'''
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import sys


def min_similarity(subset):
    '''
    Takes all entries in a dataframe column ('FundName') and compares them
    pairwise with the first entry, assigning a fuzzy string matching score, and
    tracking the minimum score for the entire group.

    Takes a DataFrame as an argument (must have 'FundName' col)
    Returns an int score representing the lowest matching score in the group
    '''
    numItems = subset.shape[0]
    base = None
    minimum = 100
    for i in xrange(numItems):
        if base is None:
            base = subset['FundName'].iloc[i]
            base = fuzz.full_process(base, force_ascii=True)
        else:
            other = subset['FundName'].iloc[i]
            other = fuzz.full_process(other, force_ascii=True)
            score = process.extractOne(other, [base, ])
            if score[1] < minimum:
                minimum = score[1]
    return minimum


if __name__ == '__main__':

    isins = pd.read_csv(sys.argv[1])
    # apply min_similarity to fund names by ISIN
    clusters = isins[['ISIN', 'FundName']].groupby(
        ['ISIN']).apply(min_similarity)
    isin_scores = pd.DataFrame(clusters, columns=['SimilarityScore'])
    isin_scores.sort(columns=['SimilarityScore'], inplace=True)
    # retain only scores at or below threshould (60)
    low_scores = isin_scores[isin_scores['SimilarityScore'] <= 60]
    low_scores.to_csv('%s%s' % (sys.argv[2], 'isin_scores.csv'))
