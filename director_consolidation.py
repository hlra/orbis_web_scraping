import pandas as pd
import os
import codecs
import csv

readpath = r"G:\ORBIS\Directors\\"
writepath = r"G:\ORBIS\Directors\\"

def stripquotes(x):
    x = str(x)[1:-1]
    return x



# next = pd.read_csv(readpath + "OWNHIST2.txt", sep="|", encoding="UTF-16", index_col=0).loc[1:]

# if os.path.exists(writepath + "directors.csv"):
#     os.remove(writepath + "directors.csv")

try:
    finished_files = list(pd.read_csv(writepath+"incl_files.csv").loc[:,1])
except:
    finished_files = []
    next = pd.read_csv(readpath + "DM149.txt", sep="|", encoding="UTF-16", usecols = ["BvD ID number"])
    next.to_csv(writepath + "directors.csv", mode="a", encoding="UTF-16", sep=";")
    next = next.iloc[1:,].drop_duplicates()
    next = pd.Series(next["BvD ID number"])
    finished_files.append("DM149.txt")
    pd.Series(finished_files).to_csv(writepath+"incl_files.csv")

for r, d, f in os.walk(readpath):
    for file in f:
        if '.txt' in file and file not in finished_files:
            print("Appending file " + str(file)+". {0:.2f}% of consolidation finished.".format((((f.index(file)+1)/len(f))*100)))
            ## Read in directors and manager file
            try:
                nextdirectors = pd.read_csv(readpath + file, sep="|", encoding="UTF-16")
                nextdirectors.to_csv(writepath+"directors.csv", mode="a", encoding="UTF-16", sep=";")
                nextdirectors = pd.Series(nextdirectors["BvD ID number"])

            except:
                try:
                    nextdirectors = pd.read_csv(readpath+file, sep="|", encoding = "UTF-16", quoting=csv.QUOTE_NONE)
                    nextdirectors.to_csv(writepath + "directors.csv", mode="a", encoding="UTF-16", sep=";")
                    nextdirectors = pd.Series(nextdirectors["BvD ID number"])
                except:
                    nextdirectors = pd.read_csv(readpath+file, sep="|", encoding = "UTF-16", header=None, quoting=csv.QUOTE_NONE, error_bad_lines=False)
                    nextdirectors = nextdirectors.applymap(stripquotes)
                    nextdirectors.columns = nextdirectors.iloc[0]
                    nextdirectors = nextdirectors.iloc[1:]
                    nextdirectors.to_csv(writepath + "directors.csv", mode="a", encoding="UTF-16", sep=";")
                    nextdirectors = pd.Series(nextdirectors["BvD ID number"])

            nextdirectors = nextdirectors.loc[1:].drop_duplicates().astype(str)

            ## Drop downloaded headers
            # nextdirectors = nextdirectors.drop(nextdirectors.index[0:2])
            ## Only keep relevant columns
            # nextdirectors = nextdirectors.filter(
            #     items=["BvD ID number"])
            # nextdirectors.drop_duplicates()
            next = next.astype(str)
            next = next.append(nextdirectors)
            finished_files.append(file)
            pd.Series(finished_files).to_csv(writepath + "incl_files.csv")

## Reset index
nextdirectors = next.reset_index(drop=True)
## Append all directors together
nextdirectors.to_csv(writepath+"directorids.csv", sep="|", encoding="UTF-16")

## Add BvD IDs
ids = pd.read_csv(r"C:\Users\Sakul\Seafile\Meine Bibliothek\Dissertation\Data\ORBIS\Scraping\Scraped_Data\\"+"full_sample_info.csv", sep="|", encoding="UTF-16", usecols=list(range(2,8)))

nextdirectors = pd.read_csv(writepath+"directorids.csv", sep="|", usecols=["BvD ID number"], encoding="UTF-16")
nextdirectors = nextdirectors.iloc[:,0].apply(lambda x : str(x)).sort_values().drop_duplicates()
nextdirectors = list(nextdirectors)

ids = ids["BvD ID number"]
ids = ids.apply(lambda x : str(x)).sort_values()
ids = list(ids.drop_duplicates())


## Which IDs are in the sample but we do not have ownership data yet?
ids_rest_own = pd.DataFrame(list(set(ids).difference(nextdirectors))).dropna().sort_values(by=0, ignore_index=True)
ids_rest_own.to_csv(writepath+"sample_rest.bvd", index = False, header = False)

## Add BvD IDs
wallenbergs = pd.read_csv(r"C:\Users\Sakul\Seafile\Meine Bibliothek\Dissertation\Data\ORBIS\Scraping\Scraped_Data\\"+"wallenbergs.csv")
wallenbergs = list(wallenbergs.iloc[:,0])
set(wallenbergs).difference(ids)
wallenbergs_stift = pd.read_csv(r"C:\Users\Sakul\Seafile\Meine Bibliothek\Dissertation\Data\ORBIS\Scraping\Scraped_Data\\"+"wallenbergs_stift.csv")
set(wallenbergs_stift).difference(ids)

# wallenbergs_group = pd.read_csv(r"C:\Users\Sakul\Desktop\\"+"Wallenberg_Corporate_Group.txt", encoding="UTF-16", sep="|")
# wallenbergs_group.to_excel(r"C:\Users\Sakul\Desktop\\"+"Wallenberg_Corporate_Group.xlsx", encoding="UTF-16")