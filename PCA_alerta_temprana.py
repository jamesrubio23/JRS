# -*- coding: utf-8 -*-
"""
Created on Mon Apr 23 12:25:05 2018

@author: 75561498
"""
# http://slendermeans.org/ml4h-ch8.html

import pandas as pd
import xlwings as xw  
import numpy as np
from sklearn.decomposition import PCA
from pandas import *
import datetime as dt
import statsmodels.api as sm
import matplotlib.pyplot as plt
from sklearn.preprocessing import StandardScaler # for standardizing the Data
from sklearn.decomposition import PCA # for PCA calculation
from statsmodels.tsa.api import VAR
import seaborn as sns



#############################################################################
################ LOADING DATA ###############################################
#############################################################################

filename = 'R:/_BankAnalytics/Centros/CoyEcono/Datos/publico/Modelizacion/Alerta_temprana/indicadores_sectoriales.xlsx'
filename = '//ofilbk01/datos/_BankAnalytics/Centros/CoyEcono/Datos/publico/Modelizacion/Alerta_temprana/indicadores_sectoriales.xlsx'
filename = 'C:/Users/u61498n/Desktop/Factor model/Indicadores_julian/indicadores_sectoriales.xlsx'
filename = 'Y:/mesa teso/JRS/1_2_PCA Ciclo mensual semanal/indicadores_sectoriales.xlsx'

Book = xw.Book(filename)
sector=['SOFT','Hard','Mercado','Inflacion','Mercado_CONDFIN_w','Mercado_RV_w','Mercado_TIPOS_w']
#sector=['Mercado_CONDFIN','Mercado_RV','Mercado_TIPOS','US_yield_Oil_USD_PMI']


n_pca = 1  # seleccionar el numero de componentes

for i in sector:
    print (i)
    
    Book.sheets(i).range("Ay15:ba4000").clear_contents()
    Book.sheets(i).range("Bn15:bz4000").clear_contents()
    
    
    df = Book.sheets(i).range('ai14')\
                .options(pd.DataFrame,expand='table',index=False).value
                
    ###### DFM
    mod = sm.tsa.DynamicFactor(endog=df, k_factors=1, factor_order=2, error_order=2)
    initial_res = mod.fit(method='powell', disp=False)
    res = mod.fit(initial_res.params, disp=False)
    print(res.summary(separate_params=False))
    Book.sheets(i).range('ba15').value = res.factors.filtered.transpose()   
    res.plot_coefficients_of_determination(figsize=(8,2));
   
    # Obtener predicciones in sample y out of sample para 24 periodos
    prediction = res.get_prediction(start=0, end=len(df) + 24)
    point_predictions = prediction.predicted_mean    
    Book.sheets(i).range('bN15').value = point_predictions   


    ###### PCA MÉTODO 1

    calc_returns = lambda x: np.log(x / x.shift(1))[1:]
    scale        = lambda x: (x - x.mean()) / x.std()
        
    def OLSreg(y, Xmat):
        return sm.OLS(y, sm.add_constant(Xmat, prepend = False)).fit()
    
    def make_pca_index(data, scale_data = True):
        '''
        Compute the correlation matrix of a set of stock data, and return
        the first principal component.
    
        By default, the data are scaled to have mean zero and variance one
        before performing the PCA.
        '''
        if scale_data:
            data_std = data.apply(scale)
        else:
            data_std = data
        corrs = np.asarray(data_std.cov())
        pca   = PCA(n_components = n_pca).fit(corrs)
        mkt_index = -scale(pca.transform(data_std))
        print(pca.explained_variance_ratio_)
        return mkt_index 
              
    pca_indicador = make_pca_index(df.dropna())
    #pca_explicada = make_pca_explained(df)
    Book.sheets(i).range('ay15').value = pca_indicador
    
    ###### PCA MÉTODO 2
    X = df.dropna().values # getting all values as a matrix of dataframe 
    sc = StandardScaler() # creating a StandardScaler object
    X_std = sc.fit_transform(X) # standardizing the data
    #X_std = df.values
    
    pca = PCA(n_pca)  
    X_pca = pca.fit_transform(X_std) # fit and reduce dimension
    Book.sheets(i).range('az15').value = X_pca
    Book.sheets(i).range('bc4').value = pca.components_
    
    print(np.cumsum(pca.explained_variance_ratio_))
    plt.plot(pca.components_)
    plt.xlabel('number of components')
    plt.ylabel('cumulative explained variance');
   
    
    """
    # We center the data and compute the sample covariance matrix.
    X_centered = X - np.mean(X, axis=0)
    n_samples = X.shape[0]
    cov_matrix = np.dot(X_centered.T, X_centered) / n_samples
    eigenvalues = pca.explained_variance_
    for eigenvalue, eigenvector in zip(eigenvalues, pca.components_):    
        print(np.dot(eigenvector.T, np.dot(cov_matrix, eigenvector)))
        print(eigenvalue)
    """
    
 
    
 
""" 
#####################################################################
################# VAR e IRF #########################################
#####################################################################

# Vamos a estimar un VAR y sacar las IRF de los indicadores
"""

# Cargar los datos de la hoja 'ResumenPCA'
sheet_name = 'ResumenPCA'
sheet_name = 'ResumenPCA_MERCADO_w'

df = Book.sheets(sheet_name).range('A14').options(pd.DataFrame, expand='table', index=True).value
# Filtrar los datos hasta diciembre de 2019
#df = df.loc['2010-12-31':]
df = df.loc[:'2019-12-31']
df = df.dropna()
#df = df[['Mercado','Encuestas','Reales','Inflación']]
#df = df[['Mercado','Encuestas','Reales']]
df = df[['Cíclicos','Tipos','Cond.Fin.']]

#df = df[['Mercado','Diff Encuestas-DatosReales']]

# Estimar el modelo VAR
model = VAR(df)
results = model.fit(maxlags=2, ic='aic')

# Imprimir resumen de los resultados
print(results.summary())

# Generar funciones de impulso respuesta (IRF)
irf = results.irf(10)

# Graficar IRF
irf.plot(figsize=(10, 8))
plt.show()

# Guardar IRF en Excel
irf_values = irf.irfs
#Book.sheets(sheet_name).range('r10').value = irf_values





