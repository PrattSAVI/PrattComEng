#%% IMPORT and CONFIG

import pandas as pd

path = r"C:\Users\csucuogl\Dropbox\ComEng_Viz\Pratt Community Partnered Course Catalogue - EDIT.xlsx"
df = pd.read_excel( path , sheet_name= 'Community Partnered Course Cata' )
df = df.dropna( axis=1, how='all')
df.head()

# %% Multiply Multi Enteries

#Many parters are noted here in a single line
dfm = df[ df['Name of Community Partner'].str.contains("\n") ]
dfm

# %% Multiple Lines and remove additional enteries

temp = pd.DataFrame()
for i,r in dfm.iterrows():

    print (r)
    num = len( r['Name of Community Partner'].split("\n") )
    print( r['Name of Community Partner'])
    print( r.str.contains("\n") )
    cols = r.index[ (r.str.contains("\n")) & (~pd.isna(r)) ].tolist()

    for n in range(num):

        r2 = r.copy()

        for c in cols:
            print(c)
            try:
                r2[c] = r2[c].split("\n")[n]
            except:
                r2[c] = None

        temp = temp.append( r2 )
temp

# %%

dfs = df[ ~ df['Name of Community Partner'].str.contains("\n") ]

df2 = dfs.append( temp )

df2.head()
# %%

pd.set_option('display.max_rows', 2000)

df2.groupby( "Name of Community Partner").size().sort_values()
# %% CHECK similar Names

from fuzzywuzzy import process, fuzz

for name in df2["Name of Community Partner"].unique():

    
    res = process.extract( name, df2[ df2["Name of Community Partner"] != name ]["Name of Community Partner"] ,
        scorer=fuzz.token_sort_ratio)

    match = pd.DataFrame( res , columns = ['name','s1','s2'])
    match = match[ match['s1'] > 75 ]
    if len( match) > 0:
        print (name)
        display( match )
        print('-------------------------')
    

# %% CHECK similar Names
pd.set_option('display.max_colwidth', None)
from fuzzywuzzy import process, fuzz

col_name = 'Pratt Department or Center affiliated with Community Partnered Course'
for name in df2[col_name].unique():

    
    res = process.extract( name, df2[ df2[col_name] != name ][col_name] ,
        scorer=fuzz.token_sort_ratio)

    match = pd.DataFrame( res , columns = ['name','s1','s2'])
    match = match[ match['s1'] > 75 ]
    if len( match) > 0:
        print (name)
        display( match )
        print('-------------------------')
    


# %%

df2[col_name] = df2[col_name].replace( "Graduate Center for Planning and the Environment – Planning","Graduate Center for Planning and the Environment – Planning")


df2.to_excel( r"C:\Users\csucuogl\Dropbox\ComEng_Viz\Pratt Community Partnered Course Catalogue - EDIT_2.xls" )
# %% Simple Network

df2['color'] = None
df2.loc[ df2['Pratt School'] == "Institutional" , "color"] = "red"
df2.loc[ df2['Pratt School'] != "Institutional" , "color"] = "blue"

import networkx as nx
import matplotlib.pyplot as plt 
G = nx.Graph()

G=nx.from_pandas_edgelist(
    df2, 
    'Pratt Department or Center affiliated with Community Partnered Course', 
    'Name of Community Partner',
    edge_attr=["color"])

plt.figure( figsize=(14,14))
nx.draw(G, with_labels=True , font_size=7 , node_size=50)
plt.show()
# %%


df.head()
# %%


df2[ df2['Pratt School'] != "Institutional" ]
# %%
