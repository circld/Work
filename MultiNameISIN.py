import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import sys


def min_similarity(subset):
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
            if score[1] < 100:
                minimum = score[1]
    return minimum


if __name__ == '__main__':

    isins = pd.read_csv(sys.argv[1])
    clusters = isins[['ISIN', 'FundName']].groupby(
        ['ISIN']).apply(min_similarity)
    isin_scores = pd.DataFrame(clusters, columns=['SimilarityScore'])
    isin_scores.sort(columns=['SimilarityScore'], inplace=True)
    import pdb; pdb.set_trace()  # XXX BREAKPOINT
    low_scores = isin_scores[isin_scores['SimilarityScore'] <= 60]
    low_scores.to_csv('%s%s' % (sys.argv[2], 'isin_scores.csv'))
