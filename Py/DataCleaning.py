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

# %% Interactive network

df2['Dept'] = df2['Pratt Department or Center affiliated with Community Partnered Course' ]
df2['Partner'] = df2['Name of Community Partner']

df2['Dept'] = df2['Dept'].str.replace('Graduate Center for Planning and the Environment',"GCPE" , regex = True)
df2['Dept'] = df2['Dept'].replace('Graduate Architecture and Urban Design',"GAUD")
df2['Dept'] = df2['Dept'].replace('Sustainable Environmental Systems',"SES" , regex = True)

df2['Dept'] = df2['Dept'].replace('Pratt Center for Community Development ','Pratt Center for Community Development')
df2['Dept'] = df2['Dept'].replace('Pratt Center for Community Development','Pratt Center')

df2['Dept'] = df2['Dept'].replace('Spatial Analysis and Visualization Initiative','SAVI')
df2['Dept'] = df2['Dept'].replace('GCPE –\xa0Planning','GCPE – Planning')

df2['Dept'] = df2['Dept'].replace('Interior Design ','Interior Design')
df2['Dept'] = df2['Dept'].replace('The Consortium for Research and Robotics ','CRR')

df2['Dept'] = df2['Dept'].replace('Center for Art, Design, and Community Engagement K-12','K-12 Center')


df2['Dept'].unique().tolist()

#%%

df2.to_excel( r"C:\Users\csucuogl\Dropbox\ComEng_Viz\Pratt Community Partnered Course Catalogue - EDIT_2.xls" )


#%%

def addFlower(mother,child):
    print(mother,child)

    [net.add_node(n, label=n) for n in child]
    net.add_node(mother, label=mother)
    [net.add_edge(mother, n, weight=1,value=2) for n in child ]

from pyvis.network import Network
net = Network(notebook=False, width="100%", height="100%")

nodes = df2['Dept'].tolist() + df2['Partner'].tolist()

#Dept - School

gcpes = [
    'GCPE – SES',
    'SES – Historic Preservation',
    'GCPE – Historic Preservation',
    'GCPE – Planning',
    'GCPE – Urban Placemaking and Management'
    ]

addFlower("GCPE" , gcpes )

centers = [
    "CRR",
    "Pratt Center",
    "SAVI",
    "K-12 Center"
]

addFlower("Centers" , centers )

#Dept - Com
[net.add_node(n, label=n) for n in nodes]
[net.add_edge(r['Dept'], r['Partner'], weight=.87) for i,r in df2.iterrows() ]

net.show(r"C:\Users\csucuogl\Documents\GitHub\PrattComEng\img\Network1.html")
# %%

df2['Dept'].unique()
# %%
