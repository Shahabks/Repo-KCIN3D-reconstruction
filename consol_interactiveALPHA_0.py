import pandas as pd
import numpy as np

prnamepv = input ( "Enter the project plot name: " )
prnamelv = input ( "Enter the project profile name: " )
querytagA = int(input("Enter the 1st tag for DX: "))
querytagB = int(input("Enter the last tag for DX: "))

mdpathpv = '/mnt/c/Users/CINPC0075/Desktop/Repo-KCIN3D-reconstruction/'+ prnamepv +'.csv'
mdpathlv = '/mnt/c/Users/CINPC0075/Desktop/Repo-KCIN3D-reconstruction/'+ prnamelv +'.csv'
dflv = pd.read_csv(mdpathlv,index_col=0)
dfpv = pd.read_csv(mdpathpv,index_col=0)

to_drop = ['Author','Closed', 'Comments', 'Drawing Revision Number', 'EdgeStyleId', 'FaceStyleId', 'File Accessed', 'File Created', 'File Last Saved By', 'File Location',
           'File Modified','File Name', 'File Size', 'Fit Smooth', 'Hyperlink', 'Hyperlink Base', 'Keywords']
dfpv.drop(to_drop, inplace=True, axis=1)
dflv.drop(to_drop, inplace=True, axis=1)

dfpv.set_index("Value",inplace= True)
dflv.set_index("Value",inplace= True)

dtag = []
ran = range(querytagA,querytagB)
for i in ran:
    m = "NO."+str(i)
    dtag.append(m)
row = dtag

dff = dfpv.loc[row, ["Position X", "Position Y", "Position Z"]]
dff.sort_values("Position X", inplace=True, ascending=True)
dff

df1 = dfpv.loc[dfpv["Name"] == 'Circle']
df2 = df1.loc[df1["Radius"] == 2.5]
df22 = df2.loc[df2["Name"] == "Circle", ["Center X", "Center Y", "Center Z"]]
df22.sort_values("Center X", inplace=True, ascending=True)
rrx = df22["Center X"]
rry = df22["Center Y"]
rrrx = dff["Position X"]
rrry = dff["Position Y"]
rrrr0 = dff.index
import math
fel=[]
for i, j in zip(rrx,rry):
    for m, n, k in zip(rrrx,rrry,rrrr0):
        f= math.sqrt((j-n)**2 + (i-m)**2)
        fel.append([i,j,m,n,k,f])
fel
from pandas import DataFrame
d1 = pd.DataFrame(fel, columns=['CX','CY','TX','TY','TG','Dis'])
d11 = d1[:30]
d110 = d11.sort_values(by='Dis', ascending=True)
d11out=d110[:1]
d12 = d1[30:60]
d120 = d12.sort_values(by='Dis', ascending=True)
d12out= d120[:1]
d13 = d1[60:90]
d130 = d13.sort_values(by='Dis', ascending=True)
d13out = d130[:1]
d14 = d1[90:120]
d140 = d14.sort_values(by='Dis', ascending=True)
d14out = d140[:1]
d15 = d1[120:150]
d150 = d15.sort_values(by='Dis', ascending=True)
d15out = d150[:1]
d16 = d1[150:180]
d160 = d16.sort_values(by='Dis', ascending=True)
d16out = d160[:1]
d17 = d1[180:210]
d170 = d17.sort_values(by='Dis', ascending=True)
d17out = d170[:1]
dframe =  [d11out,d12out,d13out,d14out,d15out,d16out,d17out]
fi = pd.concat(dframe)
fi

dl1 = dfpv.loc[dfpv["Name"] == 'Line']
dl2 = dl1.loc[dl1['Color'] == "red"]
dl3 = dl2.loc[dl2["Length"] == 5]
dl4 = dl3.loc[dl3['Name'] == "Line", ["End X", "End Y", "End Z",'Start X','Start Y','Start Z']]
dl4.reset_index(drop=True, inplace=True)
dl4.sort_values("End X", inplace=True, ascending=True)
dl4


print(fi,dl4)