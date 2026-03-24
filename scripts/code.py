import pandas as pd
import numpy as np
import geopandas as gpd
from matplotlib import pyplot as plt
import time
start_time0 = time.time()
import rasterio
from rasterio.mask import mask
import warnings
from scipy.stats import pearsonr, linregress


edgecolor = '#505050'
warnings.filterwarnings('ignore')
plt.rcParams['font.family']=['Times New Roman']
plt.rcParams['font.weight']='bold'


def r2(x,y):
    x,y = np.array(x),np.array(y)
    corr, p_value = pearsonr(x, y)
    result = linregress(x, y)
    return {'r':result.rvalue,'p':p_value,'斜率':result.slope,'截距':result.intercept}

def calculate_power_generation_offshore_wind(capacity): ## Mw
    capacity = 10**6 * capacity / (10**3)
    ## kwh
    return 8760 * (capacity*0.0295948*0.55+capacity*0.110219887*0.5+capacity*0.200110635*0.45+capacity*0.660074678*0.35)

def sub(a):##Subscript
    a=str(a)
    return '$\mathregular{_{'+a+'}}$'

def sup(a):##Superscript
    a=str(a)
    return '$\mathregular{^{'+a+'}}$'

work_path = 'D:/workplace'

'''
EFs
'''
efs = pd.read_excel(work_path+'/data/efs_origin_data.xlsx',sheet_name = None,index_col=0)
efs1,efs2 = efs['Sheet1'], efs['Sheet2'].T
efs1.columns = ['region',2023,2022,2021]
efs2.reset_index(drop=False,inplace=True)
efs2['year'] = efs2.iloc[:,0].apply(lambda x: x.split('-')[0])
efs2 = efs2.iloc[:,1:].groupby('year').mean()
for year in np.arange(2020,2005,-1):
    efs1[year] = 0
    for province in efs1.index:
        region = efs1.loc[province,'region']
        ratio = efs1.loc[province,2021]/efs2.loc['2021',region]
        if province == '海南' and int(year) in range(2006,2011):
            efs1.loc[province,year] = efs2.loc[str(year),province]
        else:
            efs1.loc[province,year] = efs2.loc[str(year),region]*ratio
efs1 = efs1.loc[:,range(2006,2024)]
for province in efs1.index:
    r = r2(range(2021,2024),efs1.loc[province,2021:2023])
    a,b = r['斜率'],r['截距']
    efs1.loc[province,2024] = a*2024+b
efs1 = efs1.loc[:,range(2006,2025)]
efs1.to_excel(work_path+'/data/efs.xlsx')

'''
MEIC data process
'''
meic_data = pd.read_excel(work_path+'/data/meic/MEIC_carbon_emission_Global_2000-2024.xlsx'
                          ,sheet_name = 'MEIC-global-CO2 by sector'
                          ,skiprows = 7)

meic_data1 = meic_data.loc[:,'Sub-country':]
meic_data1.drop('Sector',axis=1,inplace=True)
meic_data1 = meic_data1.groupby(['Sub-country','Month']).sum()
meic_data1.to_excel(work_path+'/data/meic/MEIC_carbon_emission_ALL_2000-2024.xlsx')


# 
'''
National average EFs
'''
elec_prod_data = pd.read_excel(work_path+'/data/CHINA_ELEC_PRODUCTION.xlsx',index_col = 0 ).loc[:,2006:]
efs = pd.read_excel(work_path+'/data/efs.xlsx',index_col = 0,)
efs.columns = range(2006,2025)
elec_prod_data.columns =range(2006,2025)
efs_avg = (elec_prod_data*efs).sum(axis=0)/elec_prod_data.sum(axis=0)
efs_avg.to_excel(work_path+'/Manuscript/fig/fig1-efs.xlsx')
y0 = np.array(efs_avg)
x0 = np.arange(2006,2025)
fig,ax = plt.subplots()

ax.plot(x0,y0,linewidth = 1.7,c = '#A4ABD6',marker = 'o',markersize = 4,label = 'This study')
ax.set_xticks(range(2006,2025,2))
ax.set_ylabel('EFs (kg CO'+sub(2)+'/kWh)',fontdict = {'weight':'bold','size':14})

world_efs = pd.read_excel(work_path+'/data/global_efs.xlsx',index_col = 0).loc[:2024,:]
s = 15
c = '#D24735'
ax.scatter(world_efs.index,world_efs['EU'],marker = '*',c = c,label = 'IEA-EU',s = s)
ax.scatter(world_efs.index,world_efs['US'],marker = 'v',c = c,label = 'IEA-US',s = s)
ax.scatter(world_efs.index,world_efs['India'],marker = '^',c = c,label = 'IEA-India',s = s)

#Implications of Climate Change on Wind Energy Potential
c = '#FAC03D'
ax.scatter([2022],[0.632],label = 'Ref-India',c = c,marker = '^')
ax.scatter([2022],[0.367],label = 'Ref-US',marker = 'v',c = c)
ax.scatter([2022],[0.257],label = 'Ref-UK',c = c,marker = '+')
ax.scatter([2022],[0.085],label = 'Ref-France',c = c,marker = 'H')

ax.legend()
plt.savefig(work_path+'/Manuscript/fig/fig1-efs.png',dpi=300,bbox_inches='tight')
plt.show()

'''
National carbon emission and Power sector carbon emission
'''
data0 = pd.read_excel(work_path+'/data/meic/MEIC_carbon_emission_ALL_2000-2024.xlsx')
data0.fillna(method='ffill',axis=0,inplace=True)
data1 = pd.read_excel(work_path+'/data/meic/MEIC_carbon_emission_Global_Power_generation_2000-2024.xlsx')
data1.drop('Sector',axis=1,inplace=True)

data0 = data0.groupby('Sub-country').sum()
data1 = data1.groupby('Sub-country').sum()

data0,data1 = data0.loc[:,2006:].sum(axis=0),data1.loc[:,2006:].sum(axis=0)
data = pd.DataFrame([data0,data1],index=['ALL','POWER']).T
data.to_excel(work_path+'/Manuscript/fig/fig2-emissions.xlsx')
fig,ax = plt.subplots()
x = range(2006,2025)
ax.bar(x=x,height=data0,label = 'All sectors',edgecolor='grey',color = '#5AA4AE',)
ax.bar(x=x,height=data1,bottom=data0-data1,label = 'Power generation',edgecolor='grey',color = '#F5F2E9')
ax.set_xticks(np.arange(2006,2025,2))
ax.set_ylabel('CO'+sub(2)+'reduction (Mt)',fontdict = {'weight':'bold','size':14})
ax.legend()
plt.savefig(work_path+'/Manuscript/fig/fig2-emissions.png',dpi=300,bbox_inches='tight')
plt.show()

'''
Carbon reduction from wind power
'''

efs = pd.read_excel(work_path+'/Manuscript/fig/fig1-efs.xlsx',index_col = 0).iloc[:,0]
capacity = pd.read_excel(work_path+'/data/Offshore_wind_power_capacity.xlsx',index_col=0).iloc[:2024,0]  ##MW
capacity = capacity.loc[:2024,].T
capacity = calculate_power_generation_offshore_wind(capacity)
carbon_reduction = efs*capacity ## kg
carbon_reduction = carbon_reduction / 10**9 ##Mt
carbon_reduction_cumsum = carbon_reduction.cumsum()

data = pd.DataFrame([carbon_reduction,carbon_reduction_cumsum],index=['data','cumsum']).T
data.to_excel(work_path+'/Manuscript/fig/fig3-carbon_redcution.xlsx')

world_capacity = pd.read_excel(work_path+'/data/Offshore_wind_power_capacity.xlsx',index_col=0).iloc[:2024,1:3] 
capacity_factor = {'EU':0.408, # Correlation between the Production of Electricity by Offshore Wind Farms and the Demand for Electricity in Polish Conditions
                   'world':0.33 # Platform Optimization and Cost Analysis in aFloating Offshore Wind Farm
                   }
world_capacity['EUROPE'] = 10**3* capacity_factor['EU'] * world_capacity['EUROPE'] * 8760
world_capacity['REST OF WORLD'] = 10**3*  capacity_factor['EU'] * world_capacity['REST OF WORLD'] * 8760
world_efs = pd.read_excel(work_path+'/data/global_efs.xlsx',index_col = 0)
world_carbon_reduction = pd.DataFrame(index = range(2015,2025))
world_carbon_reduction['EU'] = world_capacity.loc[2015:2024,'EUROPE']*world_efs.loc[2015:2024,'EU']
world_carbon_reduction['REST OF WORLD'] = world_capacity.loc[2015:2024,'REST OF WORLD']*world_efs.loc[2015:2024,'WORLD']
world_carbon_reduction = world_carbon_reduction / 10**9 ##Mt
world_carbon_reduction.to_excel(work_path+'/Manuscript/fig/fig3-world_carbon_redcution.xlsx')
fig,ax = plt.subplots()
x = range(2006,2025)
ax.bar(x=x,height=data.iloc[:,1],label = 'Total reduction',edgecolor='grey',color = '#EEEAD9')
ax.bar(x=x,height=data.iloc[:,0],bottom=data.iloc[:,1]-data.iloc[:,0]
       ,label = 'Additional reduction',edgecolor='grey',color = '#EEEAD9',hatch = '////')
ax.set_xticks(np.arange(2006,2025,2))
ax.set_ylabel('CO'+sub(2)+'reduction (Mt)',fontdict = {'weight':'bold','size':14})
ax.legend()

position = ax.get_position().bounds
ax = fig.add_axes([position[0]*1.7,position[1]*2.5,position[2]*0.52,position[3]*0.6])

x = np.arange(2015,2025)
width = 0.2
ax.bar(x=x-width,width = width, height = carbon_reduction.loc[2015:2024],edgecolor = 'grey',label = 'China',color = '#EEEAD9')
ax.bar(x=x,width = width,height = world_carbon_reduction['EU'],edgecolor = 'grey',label = 'EU',color = '#F29A76')
ax.bar(x=x+width,width = width,height = world_carbon_reduction['REST OF WORLD']
       ,edgecolor = 'grey',label ='Rest of World',color = '#5AA4AE')
ax.set_xticks(np.arange(2016,2026,2))
ax.legend(fontsize = 7,frameon=False,borderpad = 0.3,labelspacing = 0.1,handletextpad =0.2,loc=2)

plt.savefig(work_path+'/Manuscript/fig/fig3-carbon_redcution.png',dpi=300,bbox_inches='tight')
plt.show()

'''
Carbon reduction of provinces
'''

province_capacity = pd.read_excel(work_path+'/data/Offshore_wind_power_capacity.xlsx',sheet_name = 'Sheet2')
province_capacity = province_capacity.iloc[1,1:]
national_capacity = pd.read_excel(work_path+'/data/Offshore_wind_power_capacity.xlsx',index_col=0).iloc[:2024,0]
national_capacity = national_capacity.cumsum()
province_capacity = province_capacity * national_capacity.loc[2018]/100
province_capacity = calculate_power_generation_offshore_wind(province_capacity)
efs = pd.read_excel(work_path+'/data/efs.xlsx',index_col = 0)
efs = efs.loc[province_capacity.index,2018]
carbon_reduction = efs * province_capacity/10**9
carbon_reduction = carbon_reduction * 1000 # kt

carbon_reduction.to_excel(work_path+'/Manuscript/fig/fig4-carbon_distribution.xlsx')

pop_data = rasterio.open(work_path+'/data/tif/2018.tif')

datax = pop_data.read(1)*0
geodata0 = gpd.read_file(work_path+'/data/shp/市级.shp')
data0 = geodata0.iloc[:,:-1]

carbon_reduction0  = pd.read_excel(work_path+'/Manuscript/fig/fig4-carbon_distribution.xlsx',index_col = 0)
provinces = ['辽宁省','河北省','天津市','山东省','江苏省','上海市','浙江省','福建省','广东省','广西壮族自治区','海南省']
for province in provinces:
    
    
    geodata = geodata0[geodata0['F4'] == province]
    
    carbon_reduction = carbon_reduction0.loc[province[:2],0]
    
    ar,af = mask(pop_data,geodata['geometry']
                 ,nodata = None
                 ,filled = True)
    data1 = ar[0]
    dataz = ar[0]
    data1 = carbon_reduction * data1 / data1.sum()
    datax = np.where(data1 > 0, data1, datax)

datax *= 1000  ### t
NODATA_VALUE = -9999  # 
datax[datax == 0] = NODATA_VALUE
meta = pop_data.meta.copy()
# tiff data process
meta.update({
    'driver': 'GTiff',
    'dtype': datax.dtype,
    'count': 1,  
    'compress': 'lzw',  
    'nodata': NODATA_VALUE 
})
output_path = work_path+'/data/tif/data.tif'

with rasterio.open(output_path, 'w', **meta) as dst:
    dst.write(datax, 1)
pop_data.close()


'''
Future scenario prediction
'''
from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import LinearRegression
from sklearn.pipeline import Pipeline
from sklearn.metrics import r2_score
capacity = pd.read_excel(work_path+'/data/Offshore_wind_power_capacity.xlsx',index_col = 0).iloc[:,0]
capacity = capacity.cumsum()
capacity = calculate_power_generation_offshore_wind(capacity)
efs = pd.read_excel(work_path+'/Manuscript/fig/fig1-efs.xlsx',index_col = 0).iloc[:,0]
gdp = pd.read_excel(work_path+'/data/CHINA_GDP.xlsx',index_col =0 ).sum()
gdp_growth_rate = dict(zip(range(2025,2035)
                   ,[4.8,4.2,4.2,4,3.7,3.4,
                     3.4,3.4,3.4,3.4,]))
###data from https://www.imf.org/external/datamapper/NGDP_RPCH@WEO/OEMDC/ADVEC/WEOWORLD/CHN
df = pd.DataFrame([efs,gdp]
                  ,index=['efs','gdp'],columns = efs.index
                  ).T
df.reset_index(drop=False,inplace=True)
df.columns = ['Year','Y','X']
X = df[['X']].values  # GDP
y = df['Y'].values    # emission factors
degree = 3
model = Pipeline([
    ('poly', PolynomialFeatures(degree=degree, include_bias=False)),
    ('linear', LinearRegression())
])

model.fit(X, y)
y_pred_train = model.predict(X)
R2 = r2_score(y, y_pred_train)


gdp_future = gdp[2024]
future_years = np.arange(2025, 2035)
future_gdp = []
for i in range(2025,2035):
    gdp_future = gdp_future * (100+gdp_growth_rate[i])/100
    future_gdp.append(gdp_future)
future_gdp = np.array(future_gdp)
future_Y_pred = model.predict(future_gdp.reshape(-1, 1))
efs = pd.DataFrame({'Year': future_years,'Predicted_Carbon_Intensity': future_Y_pred})
efs.set_index('Year',drop=True,inplace=True)
efs = efs.iloc[:,0]
capacity = capacity.loc[2025:]
carbon_reduction = efs*capacity ## kg
carbon_reduction = carbon_reduction / 10**9 ##Mt
carbon_reduction = pd.DataFrame(carbon_reduction,columns = ['China'])

fig,ax = plt.subplots()
ax.bar(x = range(2025,2035),height = carbon_reduction['China'],edgecolor='grey',color = '#FAC03D')
ax.set_xticks(range(2025,2035))
ax.set_yticks(range(0,450,100))
ax.set_ylabel('CO'+sub(2)+'reduction (Mt)',fontdict = {'weight':'bold','size':14})

capacity_factor = {'EU':0.408, # Correlation between the Production of Electricity by Offshore Wind Farms and the Demand for Electricity in Polish Conditions
                   'world':0.33 # Platform Optimization and Cost Analysis in aFloating Offshore Wind Farm
                   }
capacity = pd.read_excel(work_path+'/data/Offshore_wind_power_capacity.xlsx',index_col = 0)
capacity['CHINA'] = calculate_power_generation_offshore_wind(capacity['CHINA'])
capacity['EUROPE'] = 10**3 * capacity['EUROPE'] *8760 *capacity_factor['EU']
capacity['REST OF WORLD'] = 10**3 * capacity['REST OF WORLD'] *8760 *capacity_factor['world']
capacity['ALL'] = capacity['CHINA'] + capacity['EUROPE'] + capacity['REST OF WORLD']
capacity = capacity.iloc[:,:-1]

efs = pd.read_excel(work_path+'/data/global_efs.xlsx',index_col = 0)
x_pred = np.array(list(range(2006,2015))+list(range(2028,2035)))
ls = []
for i in efs.columns:
    x = efs.index
    y = efs[i]
    R2 = r2(x,y)
    a,b = R2['斜率'],R2['截距']
    y_pred = x_pred * a + b
    ls.append(y_pred)

efs_new = pd.DataFrame(ls,columns = x_pred, index = efs.columns).T
efs = pd.concat([efs,efs_new])
efs = efs.loc[range(2006,2035),:]

carbon_reduction['EU'] = efs['EU'] * capacity['EUROPE']
carbon_reduction['Rest of World'] = efs['WORLD'] * capacity['REST OF WORLD']
carbon_reduction = carbon_reduction.loc[2025:,:] / 10**9 ##Mt
carbon_reduction.to_excel(work_path+'/Manuscript/fig/fig5-future_reduction.xlsx')

plt.savefig(work_path+'/Manuscript/fig/fig5-future_reduction.png',dpi=300,bbox_inches='tight')
plt.show()





'''
'''
stop_time0 = time.time()
print('程序运行时间为：'+str(stop_time0-start_time0) +'秒')