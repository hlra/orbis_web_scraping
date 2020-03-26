import pandas as pd
import os
import codecs
import csv

readpath = r"G:\ORBIS\Ownership\\"
writepath = r"G:\ORBIS\Ownership\\"

# next = pd.read_csv(readpath + "OWNHIST2.txt", sep="|", encoding="UTF-16", index_col=0).loc[1:]

if os.path.exists(writepath + "ownership.csv"):
    os.remove(writepath + "ownership.csv")

finished_files = []
for r, d, f in os.walk(readpath):
    for file in f:
        #if '.txt' in file and "OWNHIST1.txt" not in file:
        if '.txt' in file and file not in finished_files:
            print("Appending file " + str(file)+". {0:.2f}% of consolidation finished.".format((((f.index(file)+1)/len(f))*100)))
            ## Read in directors and manager file
            try:
                next = pd.read_csv(readpath+file, sep="|", encoding = "UTF-16", header = 0, index_col=0, usecols=list(range(0,40)), quoting=csv.QUOTE_NONE, )[1:]
            except:
                # with codecs.open(readpath+file, "r", "utf-16") as sourceFile:
                #     with codecs.open((readpath + file).replace(".txt", "") + "NEWENC.txt", "w", "latin1") as targetFile:
                #         while True:
                #             contents = sourceFile.read()
                #             if not contents:
                #                 break
                #             targetFile.write(contents)
                # next = pd.read_csv((readpath + file).replace(".txt", "") + "NEWENC.txt", sep="|", encoding="UTF-16", header=0, index_col=0,
                #                    usecols=list(range(0, 40)))[1:]
                next = pd.read_csv(readpath+file, sep="|", encoding = "UTF-16", header = 0, index_col=0, usecols=list(range(0,40)), quoting=csv.QUOTE_NONE, )[1:]
                pass

            # next = next.drop_duplicates(subset=("BvD ID number"),ignore_index=True)
            next = next.reset_index()

            ## Append all directors together
            next.to_csv(writepath + "ownership.csv", sep="|", mode="a", encoding="UTF-16")
            finished_files.append(file)

## Add BvD IDs
ids = pd.read_csv(r"C:\Users\Sakul\Seafile\Meine Bibliothek\Dissertation\Data\ORBIS\Scraping\Scraped_Data\\"+"full_sample_info.csv", sep="|", encoding="UTF-16", usecols=list(range(2,8)))

compnames = pd.Series()
for chunk in pd.read_csv(writepath + "ownership.csv", sep="|", encoding="UTF-16", chunksize = 10000, usecols=['"BvD ID number"']):
    compnames = compnames.append(chunk.drop_duplicates())

compnames = compnames.iloc[:,1]
compnames = compnames.apply(lambda x : str(x).replace("\"", ""))

# compnames = pd.DataFrame(compnames).rename(columns={'"Company name"':"Company name"})
# compnames = compnames.merge(ids[["Company name", "BvD ID number"]].drop_duplicates(subset=("BvD ID number")), on='Company name', how='left')

compnames = list(compnames)
ids = ids["BvD ID number"]
ids = list(ids.drop_duplicates())

## Which IDs are in the sample but we do not have ownership data yet?
ids_rest_own = pd.DataFrame(list(set(ids).difference(compnames))).dropna().sort_values(by=0, ignore_index=True)
ids_rest_own.to_csv(writepath+"sample_rest.bvd", index = False, header = False)

## Add BvD IDs
wallenbergs = pd.read_csv(r"C:\Users\Sakul\Seafile\Meine Bibliothek\Dissertation\Data\ORBIS\Scraping\Scraped_Data\\"+"wallenbergs.csv")
wallenbergs = list(wallenbergs.iloc[:,0])
set(wallenbergs).difference(ids)
wallenbergs_stift = pd.read_csv(r"C:\Users\Sakul\Seafile\Meine Bibliothek\Dissertation\Data\ORBIS\Scraping\Scraped_Data\\"+"wallenbergs_stift.csv")
set(wallenbergs_stift).difference(ids)

# wallenbergs_group = pd.read_csv(r"C:\Users\Sakul\Desktop\\"+"Wallenberg_Corporate_Group.txt", encoding="UTF-16", sep="|")
# wallenbergs_group.to_excel(r"C:\Users\Sakul\Desktop\\"+"Wallenberg_Corporate_Group.xlsx", encoding="UTF-16")