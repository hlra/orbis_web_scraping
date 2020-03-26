import pandas as pd
import os

readpath = r"S:\Meine Bibliotheken\Meine Bibliothek\Dissertation\Data\ORBIS\Scraping\Scraped_Data\Sample_Merge\\"
writepath = r"S:\Meine Bibliotheken\Meine Bibliothek\Dissertation\Data\ORBIS\Scraping\Scraped_Data\Sample_Merge\\"

next = pd.read_csv(readpath + "BASICINFO1.txt", sep="|", encoding="UTF-16", index_col=0).loc[1:]
next = next.drop_duplicates(subset=("BvD ID number"))

for r, d, f in os.walk(readpath):
    for file in f:
        if '.txt' in file and "BASICINFO1.txt" not in file:
            print("Appending file "+ str(f.index(file)+1) + " of " + str(len(f)) + ".")
            ## Read in directors and manager file
            next = next.append(pd.read_csv(readpath+file, sep="|", encoding="utf-16", index_col=0)[2:])

#next = next.drop_duplicates(subset=("BvD ID number"),ignore_index=True)
next = next.reset_index()

## Append all directors together
next.to_csv(writepath+"full_sample_info.csv", sep="|", mode="w+", encoding="UTF-16")

sample_info = pd.read_csv(writepath+"full_sample_info.csv", sep="|", encoding="UTF-16")

ids = sample_info["BvD ID number"].drop_duplicates()
small_sample = pd.read_csv(writepath+"small_sample.csv", header = None).drop_duplicates()[0]

ids_rest = pd.DataFrame(list(set(ids).difference(small_sample))).dropna().sort_values(by=0, ignore_index=True)
# ids_rest =  [item for item in list(ids) if item not in list(small_sample)]

ids_rest1 = ids_rest.iloc[0:1000000].to_csv(writepath+"full_sample1.bvd", index = False, header = False, encoding="UTF-16")
ids_rest2 = ids_rest.iloc[1000000:2000000].to_csv(writepath+"full_sample2.bvd", index = False, header = False, encoding="UTF-16")
ids_rest3 = ids_rest.iloc[2000000:].to_csv(writepath+"full_sample3.bvd", index = False, header = False, encoding="UTF-16")

ids_rest.to_csv(writepath+"full_sample.bvd", index = False, header = False, encoding="UTF-16")
ids_rest.to_csv(writepath+"full_sample.bvd", index = False, header = False)

richmatches1 = pd.read_csv(writepath+"RichMatches0-1000.csv", header =None)
## Zero of the first 1000 firms related to a super-rich individual are not indluded in a sample
check_rich_sample = list(set(richmatches1).difference(ids))